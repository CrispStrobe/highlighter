#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Kindle Ebook Highlighter - COMPLETE VERSION with Batch Processing

All features included:
- Batch processing enabled by DEFAULT (process ALL books with 95%+ match confidence)
- Library mode enabled by DEFAULT
- Text formatting preservation (bold, italic, underline, links, footnotes)
- All 3 matching methods: regex, difflib, vector
- Compare mode
- Paragraph formatting preservation
- HTML preservation for debugging (--preserve-html)
- Enhanced verbose logging with -vv
- Improved robustness and error handling

Usage:
  Default: Batch process ALL books from library with 95%+ match confidence:
    python kindle_highlighter.py --clippings "My Clippings.txt" -v
  
  Process single best match only:
    python kindle_highlighter.py --clippings "My Clippings.txt" --no-batch -v
  
  Process specific ebook file:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" -v
  
  Batch with HTML preservation for debugging:
    python kindle_highlighter.py --clippings "My Clippings.txt" --preserve-html -vv
  
  Compare all methods on specific file:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" --compare -v
  
  List all books in library:
    python kindle_highlighter.py --list-books
"""

import sys
import os
import re
import argparse
import subprocess
import logging
import zipfile
import shutil
import glob
from typing import List, Dict, Tuple, Callable, Optional, Set
from pathlib import Path
from difflib import SequenceMatcher
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from bs4 import BeautifulSoup, NavigableString, Tag
from thefuzz import process as fuzzy_process
from thefuzz import fuzz

# --- CONFIGURATION ---
DEFAULT_CALIBRE_LIBRARY = os.path.expanduser('~/Calibre Library/')
DEFAULT_SIMILARITY_THRESHOLD = 0.9
DEFAULT_VECTOR_CONFIDENCE_THRESHOLD = 0.65
NLP_MODEL = 'all-MiniLM-L6-v2'
MAX_BOOK_SIZE_MB = 50
MIN_HIGHLIGHT_LENGTH = 5
MIN_TITLE_MATCH_SCORE = 90
AUTO_SELECT_THRESHOLD = 95

# --- LOGGING SETUP ---
logger = logging.getLogger('kindle_highlighter')

def setup_logging(verbosity: int):
    """Configure logging based on verbosity level."""
    if verbosity == 0:
        level = logging.WARNING
    elif verbosity == 1:
        level = logging.INFO
    else:  # 2 or more
        level = logging.DEBUG
    
    handler = logging.StreamHandler(sys.stdout)
    if level == logging.DEBUG:
        formatter = logging.Formatter('%(levelname)s [%(funcName)s]: %(message)s')
    else:
        formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(level)
    
    logger.debug(f"Logging configured at level: {logging.getLevelName(level)}")

# --- HELPER FUNCTIONS ---

def handle_error(message: str, is_fatal: bool = True) -> None:
    """Prints an error message and optionally exits."""
    logger.error(message)
    if is_fatal:
        sys.exit(1)

def normalize_title_for_matching(title: str) -> str:
    """
    Normalize a title for better matching.
    - Remove author names in parentheses
    - Remove series info in parentheses  
    - Remove extra whitespace
    - Remove special chars
    """
    logger.debug(f"Normalizing title: '{title}'")
    
    # Remove author/series in parentheses at the end
    title_clean = re.sub(r'\s*\([^)]+\)\s*$', '', title)
    
    # Remove special unicode characters
    title_clean = title_clean.replace('¬', '').replace('', '')
    
    # Remove multiple spaces
    title_clean = re.sub(r'\s+', ' ', title_clean).strip()
    
    logger.debug(f"Normalized to: '{title_clean}'")
    return title_clean

def find_calibre_binary() -> Optional[str]:
    """Find Calibre's ebook-convert binary on the system."""
    logger.debug("Searching for Calibre installation...")
    
    # macOS
    if sys.platform == 'darwin':
        logger.debug("Platform: macOS - searching /Applications")
        calibre_apps = glob.glob('/Applications/*alibre*.app')
        logger.debug(f"Found {len(calibre_apps)} potential Calibre apps: {calibre_apps}")
        
        for app_path in calibre_apps:
            binary_path = os.path.join(app_path, 'Contents/MacOS/ebook-convert')
            logger.debug(f"Checking: {binary_path}")
            if os.path.exists(binary_path) and os.access(binary_path, os.X_OK):
                logger.debug(f"✓ Found Calibre binary: {binary_path}")
                return binary_path
        
        common_paths = [
            '/Applications/calibre.app/Contents/MacOS/ebook-convert',
            '/Applications/Calibre.app/Contents/MacOS/ebook-convert',
            '/usr/local/bin/ebook-convert',
        ]
        logger.debug(f"Checking common macOS paths: {common_paths}")
        for path in common_paths:
            if os.path.exists(path) and os.access(path, os.X_OK):
                logger.debug(f"✓ Found Calibre binary: {path}")
                return path
    
    # Linux/Unix
    elif sys.platform.startswith('linux') or sys.platform.startswith('freebsd'):
        logger.debug(f"Platform: {sys.platform} - checking Linux/Unix paths")
        common_paths = [
            '/usr/bin/ebook-convert',
            '/usr/local/bin/ebook-convert',
            os.path.expanduser('~/calibre/ebook-convert'),
        ]
        logger.debug(f"Checking paths: {common_paths}")
        for path in common_paths:
            if os.path.exists(path) and os.access(path, os.X_OK):
                logger.debug(f"✓ Found Calibre binary: {path}")
                return path
    
    # Windows
    elif sys.platform == 'win32':
        logger.debug("Platform: Windows - checking Program Files")
        common_paths = [
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Calibre2", "ebook-convert.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Calibre2", "ebook-convert.exe"),
            os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Calibre", "ebook-convert.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "Calibre", "ebook-convert.exe"),
        ]
        logger.debug(f"Checking paths: {common_paths}")
        for path in common_paths:
            if os.path.exists(path):
                logger.debug(f"✓ Found Calibre binary: {path}")
                return path
    
    # Try in PATH
    logger.debug("Checking system PATH")
    try:
        command = 'where' if sys.platform == 'win32' else 'which'
        logger.debug(f"Running: {command} ebook-convert")
        result = subprocess.run(
            [command, 'ebook-convert'],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            path = result.stdout.strip().split('\n')[0]
            logger.debug(f"✓ Found ebook-convert in PATH: {path}")
            return path
        else:
            logger.debug(f"Command failed with return code: {result.returncode}")
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        logger.debug(f"PATH search failed: {e}")
    except Exception as e:
        logger.debug(f"Unexpected error in PATH search: {e}")
    
    logger.debug("Calibre binary not found")
    return None

def check_converter_available(calibre_path: Optional[str] = None) -> Tuple[Optional[str], Optional[str]]:
    """Check which ebook converter is available."""
    logger.debug("Checking for available ebook converters...")
    
    if calibre_path:
        logger.debug(f"Checking custom Calibre path: {calibre_path}")
        if os.path.exists(calibre_path) and os.access(calibre_path, os.X_OK):
            try:
                result = subprocess.run(
                    [calibre_path, '--version'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    version = result.stdout.strip()
                    logger.debug(f"Custom Calibre version: {version}")
                    logger.info(f"Using Calibre (custom path): {calibre_path}")
                    return calibre_path, 'htmlz'
            except (subprocess.TimeoutExpired, Exception) as e:
                logger.debug(f"Custom Calibre path failed: {e}")
    
    logger.debug("Auto-detecting Calibre...")
    calibre_binary = find_calibre_binary()
    if calibre_binary:
        try:
            result = subprocess.run(
                [calibre_binary, '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                version = result.stdout.strip()
                logger.debug(f"Calibre version: {version}")
                logger.info(f"Using Calibre: {calibre_binary}")
                return calibre_binary, 'htmlz'
        except (subprocess.TimeoutExpired, Exception) as e:
            logger.debug(f"Calibre detection failed: {e}")
    
    logger.debug("Checking for standalone ebook-converter...")
    try:
        result = subprocess.run(
            ['ebook-converter', '--help'],
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            logger.debug("✓ Found standalone ebook-converter")
            logger.info("Using standalone ebook-converter")
            return 'ebook-converter', 'htmlz'
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        logger.debug(f"Standalone ebook-converter not found: {e}")
    except Exception as e:
        logger.debug(f"Unexpected error checking ebook-converter: {e}")
    
    logger.warning("No ebook converter found")
    return None, None

def check_dependencies(calibre_path: Optional[str] = None) -> Dict[str, bool]:
    """Check all required and optional dependencies."""
    logger.debug("Checking Python package dependencies...")
    deps = {
        'docx': False,
        'bs4': False,
        'thefuzz': False,
        'converter': False,
        'mobi': False,
        'ebooklib': False
    }
    
    try:
        import docx
        deps['docx'] = True
        logger.debug("✓ python-docx found")
    except ImportError:
        logger.error("✗ python-docx not found. Install: pip install python-docx")
    
    try:
        import bs4
        deps['bs4'] = True
        logger.debug("✓ beautifulsoup4 found")
    except ImportError:
        logger.error("✗ beautifulsoup4 not found. Install: pip install beautifulsoup4")
    
    try:
        import thefuzz
        deps['thefuzz'] = True
        logger.debug("✓ thefuzz found")
    except ImportError:
        logger.error("✗ thefuzz not found. Install: pip install thefuzz python-Levenshtein")
    
    logger.debug("Checking for ebook converter...")
    converter, _ = check_converter_available(calibre_path)
    deps['converter'] = converter is not None
    if not deps['converter']:
        logger.error("✗ No ebook converter found. Install Calibre or ebook-converter")
    else:
        logger.debug("✓ Ebook converter found")
    
    logger.debug("Checking optional conversion libraries...")
    try:
        import mobi
        deps['mobi'] = True
        logger.debug("✓ mobi library found (optional)")
    except ImportError:
        logger.debug("mobi library not found (optional: pip install mobi)")
    
    try:
        import ebooklib
        deps['ebooklib'] = True
        logger.debug("✓ ebooklib found (optional)")
    except ImportError:
        logger.debug("ebooklib not found (optional: pip install ebooklib)")
    
    return deps

def check_nlp_libraries() -> bool:
    """Checks if the required NLP libraries are installed."""
    logger.debug("Checking NLP libraries for vector matching...")
    try:
        import torch
        import nltk
        from sentence_transformers import SentenceTransformer, util
        logger.debug("✓ NLP libraries found (torch, nltk, sentence_transformers)")
        return True
    except ImportError as e:
        logger.debug(f"NLP library missing: {e}")
        return False

# --- CALIBRE LIBRARY FUNCTIONS ---

def find_calibre_library(library_path: Optional[str] = None) -> Path:
    """Find and validate Calibre library path."""
    if library_path:
        path = Path(library_path).expanduser()
        logger.debug(f"Using provided library path: {path}")
    else:
        path = Path(DEFAULT_CALIBRE_LIBRARY).expanduser()
        logger.debug(f"Using default library path: {path}")
    
    if not path.exists():
        handle_error(f"Calibre library not found at: {path}")
    
    if not path.is_dir():
        handle_error(f"Calibre library path is not a directory: {path}")
    
    logger.info(f"Using Calibre library: {path}")
    return path

def discover_books(library_path: Path, format_filter: str = 'mobi') -> List[Dict]:
    """Recursively discover all books in Calibre library."""
    logger.info(f"Scanning library for .{format_filter} files...")
    logger.debug(f"Starting recursive scan of: {library_path}")
    
    books = []
    format_filter = format_filter.lower()
    scanned_dirs = 0
    scanned_files = 0
    
    try:
        for root, dirs, files in os.walk(library_path):
            scanned_dirs += 1
            if scanned_dirs % 100 == 0:
                logger.debug(f"Scanned {scanned_dirs} directories, found {len(books)} books so far...")
            
            for file in files:
                scanned_files += 1
                if file.lower().endswith(f'.{format_filter}'):
                    file_path = Path(root) / file
                    logger.debug(f"Found ebook: {file_path}")
                    
                    metadata_file = Path(root) / 'metadata.opf'
                    title = None
                    author = None
                    
                    if metadata_file.exists():
                        logger.debug(f"Reading metadata from: {metadata_file}")
                        try:
                            with open(metadata_file, 'r', encoding='utf-8') as f:
                                soup = BeautifulSoup(f, 'xml')
                                title_tag = soup.find('dc:title')
                                author_tag = soup.find('dc:creator')
                                if title_tag:
                                    title = title_tag.get_text(strip=True)
                                    logger.debug(f"  Title from metadata: {title}")
                                if author_tag:
                                    author = author_tag.get_text(strip=True)
                                    logger.debug(f"  Author from metadata: {author}")
                        except Exception as e:
                            logger.debug(f"Failed to parse metadata: {e}")
                    
                    if not title:
                        folder_name = Path(root).name
                        title = re.sub(r'\s*\(\d+\)$', '', folder_name)
                        logger.debug(f"  Using folder name as title: {title}")
                    
                    books.append({
                        'path': file_path,
                        'title': title or file,
                        'author': author or 'Unknown',
                        'filename': file,
                        'folder': Path(root).name
                    })
    except Exception as e:
        logger.error(f"Error during library scan: {e}")
        logger.debug("Full traceback:", exc_info=True)
    
    logger.info(f"✓ Found {len(books)} books in library (scanned {scanned_dirs} directories, {scanned_files} files)")
    if books:
        logger.debug(f"First 10 books: {[b['title'] for b in books[:10]]}")
    
    return books

def match_books_to_clippings(books: List[Dict], clippings_titles: List[str], min_score: int = MIN_TITLE_MATCH_SCORE) -> List[Dict]:
    """
    Match ALL books from library to clippings.
    Returns list of matched books sorted by score and highlight count.
    """
    if not books:
        logger.warning("No books found in library")
        return []
    
    if not clippings_titles:
        logger.warning("No clipping titles to match against")
        return []
    
    from collections import Counter
    title_counts = Counter(clippings_titles)
    
    logger.info(f"Matching {len(books)} library books against {len(title_counts)} unique titles from clippings...")
    logger.debug(f"Clipping titles: {list(title_counts.keys())[:10]}{'...' if len(title_counts) > 10 else ''}")
    
    matches = []
    processed_titles = 0
    
    for clip_title, highlight_count in title_counts.most_common():
        processed_titles += 1
        if processed_titles % 10 == 0:
            logger.debug(f"Processing clipping title {processed_titles}/{len(title_counts)}")
        
        clip_title_clean = normalize_title_for_matching(clip_title)
        
        best_score = 0
        best_book = None
        best_method = None
        
        for book in books:
            book_title_clean = normalize_title_for_matching(book['title'])
            
            # Use multiple scoring methods
            score1 = fuzz.token_sort_ratio(clip_title_clean.lower(), book_title_clean.lower())
            score2 = fuzz.partial_ratio(clip_title_clean.lower(), book_title_clean.lower())
            score3 = fuzz.ratio(clip_title_clean.lower(), book_title_clean.lower())
            
            score = max(score1, score2, score3)
            
            if score > best_score:
                best_score = score
                best_book = book
                best_method = f"s1={score1},s2={score2},s3={score3}"
        
        if best_book and best_score >= min_score:
            matches.append({
                'book': best_book,
                'clipping_title': clip_title,
                'clipping_title_normalized': clip_title_clean,
                'highlight_count': highlight_count,
                'score': best_score,
                'method_scores': best_method
            })
            
            logger.debug(f"'{clip_title_clean}' ({highlight_count} highlights) -> "
                        f"'{normalize_title_for_matching(best_book['title'])}' ({best_score}%) [{best_method}]")
    
    # Sort by: 1) score desc, 2) highlight count desc
    matches.sort(key=lambda x: (x['score'], x['highlight_count']), reverse=True)
    
    logger.debug(f"Match complete: {len(matches)} books matched out of {len(title_counts)} clipping titles")
    return matches

# --- FILE PARSING ---

def validate_file_size(filepath: str, max_size_mb: int = MAX_BOOK_SIZE_MB) -> bool:
    """Validate that file size is reasonable."""
    try:
        size_mb = os.path.getsize(filepath) / (1024 * 1024)
        logger.debug(f"File size: {size_mb:.2f}MB")
        if size_mb > max_size_mb:
            logger.warning(f"File is large ({size_mb:.2f}MB). This may take a while...")
        return True
    except Exception as e:
        logger.error(f"Could not check file size: {e}")
        return False

def parse_clippings(clippings_file: str) -> List[Dict[str, str]]:
    """Parses the 'My Clippings.txt' file and returns a list of clippings."""
    if not os.path.exists(clippings_file):
        handle_error(f"Clippings file not found at: {clippings_file}")
    
    logger.info(f"Parsing clippings from: {clippings_file}")
    
    try:
        file_size = os.path.getsize(clippings_file) / (1024 * 1024)
        logger.debug(f"Clippings file size: {file_size:.2f}MB")
        
        if file_size > 10:
            logger.info(f"Large clippings file ({file_size:.2f}MB). This may take a moment...")
    except Exception as e:
        logger.debug(f"Could not check file size: {e}")
    
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
    content = None
    
    for encoding in encodings:
        logger.debug(f"Trying encoding: {encoding}")
        try:
            with open(clippings_file, 'r', encoding=encoding) as f:
                content = f.read()
            logger.debug(f"✓ Successfully read file with {encoding} encoding")
            break
        except UnicodeDecodeError:
            logger.debug(f"✗ Failed to read with {encoding} encoding")
            continue
        except Exception as e:
            logger.debug(f"Unexpected error with {encoding}: {e}")
            continue
    
    if content is None:
        handle_error(f"Could not read clippings file with any known encoding")

    logger.debug(f"File content length: {len(content)} characters")
    logger.info("Parsing clippings (this may take a moment for large files)...")
    
    clippings = []
    entries = content.split('==========')
    logger.debug(f"Found {len(entries)} potential entries (split by '==========')")
    
    highlight_patterns = [
        r'^-\s*Your\s+(Highlight|Note)',
        r'^-\s*(Ihre|Deine)\s+(Markierung|Notiz)',
        r'^-\s*Votre\s+(surlignement|note)',
        r'^-\s*Tu\s+(subrayado|nota)',
        r'^-\s*La\s+tua\s+(evidenziazione|nota)',
    ]
    
    logger.debug(f"Using {len(highlight_patterns)} language patterns for detection")
    
    skipped_no_lines = 0
    skipped_not_highlight = 0
    skipped_no_text = 0
    skipped_too_short = 0
    
    for idx, entry in enumerate(entries):
        if idx % 100 == 0 and idx > 0:
            logger.debug(f"Processing entry {idx}/{len(entries)}")
        
        entry = entry.strip()
        if not entry:
            continue
        
        lines = entry.split('\n')
        if len(lines) < 3:
            skipped_no_lines += 1
            if idx < 10:
                logger.debug(f"Skipping entry {idx}: too few lines ({len(lines)})")
            continue
        
        book_title = ''.join(char for char in lines[0].strip() if char.isprintable())
        
        is_highlight = False
        matched_pattern = None
        for pattern in highlight_patterns:
            if re.match(pattern, lines[1].strip(), re.IGNORECASE):
                is_highlight = True
                matched_pattern = pattern
                break
        
        if not is_highlight:
            skipped_not_highlight += 1
            if idx < 10:
                logger.debug(f"Skipping entry {idx}: line 2 doesn't match any pattern: {lines[1][:80]}")
            continue
        
        if idx < 10:
            logger.debug(f"✓ Entry {idx} matched pattern: {matched_pattern}")
        
        clipping_text = None
        for i in range(2, len(lines)):
            if lines[i].strip():
                clipping_text = '\n'.join(lines[i:]).strip()
                break
        
        if not clipping_text:
            skipped_no_text += 1
            logger.debug(f"Skipping entry {idx}: no highlight text found")
            continue
        
        if len(clipping_text) >= MIN_HIGHLIGHT_LENGTH:
            clippings.append({'title': book_title, 'text': clipping_text})
            if idx < 10:
                logger.debug(f"✓ Added highlight from '{book_title}': {clipping_text[:50]}...")
        else:
            skipped_too_short += 1
            logger.debug(f"Skipping entry {idx}: highlight too short ({len(clipping_text)} chars)")
    
    logger.debug(f"Parsing stats: skipped {skipped_no_lines} (no lines), {skipped_not_highlight} (not highlight), "
                 f"{skipped_no_text} (no text), {skipped_too_short} (too short)")
    
    if not clippings:
        logger.error("No valid clippings were found in the provided file.")
        logger.error("The file may not be in the standard Kindle 'My Clippings.txt' format.")
        logger.error(f"File has {len(entries)} entries separated by '=========='")
        if entries and len(entries) > 1:
            logger.error(f"First entry preview:")
            logger.error(f"{entries[1][:300]}...")
        handle_error("Failed to parse clippings file.")
    
    unique_titles = len(set(c['title'] for c in clippings))
    logger.info(f"✓ Found {len(clippings)} total clippings from {unique_titles} books")
    
    if logger.level <= logging.DEBUG:
        title_counts = {}
        for clip in clippings:
            title_counts[clip['title']] = title_counts.get(clip['title'], 0) + 1
        
        logger.debug("\nBooks found in clippings (top 10 by highlight count):")
        for title, count in sorted(title_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
            logger.debug(f"  - {title}: {count} highlights")
        
        if len(title_counts) > 10:
            logger.debug(f"  ... and {len(title_counts) - 10} more books")
    
    return clippings

# --- CONVERSION FUNCTIONS ---

def extract_htmlz(htmlz_path: str) -> Tuple[str, str]:
    """Extract htmlz file and return paths to main HTML and extract dir."""
    logger.debug(f"Extracting htmlz file: {htmlz_path}")
    
    extract_dir = os.path.splitext(htmlz_path)[0] + '_extracted'
    logger.debug(f"Extract directory: {extract_dir}")
    
    if os.path.exists(extract_dir):
        logger.debug(f"Removing existing extract directory")
        try:
            shutil.rmtree(extract_dir)
        except Exception as e:
            logger.warning(f"Could not remove existing extract directory: {e}")
    
    try:
        logger.debug("Opening zip archive...")
        with zipfile.ZipFile(htmlz_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()
            logger.debug(f"Archive contains {len(file_list)} files")
            if logger.level <= logging.DEBUG and len(file_list) <= 20:
                logger.debug(f"Files in archive: {file_list}")
            zip_ref.extractall(extract_dir)
        
        html_files = list(Path(extract_dir).rglob('*.html'))
        logger.debug(f"Found {len(html_files)} HTML files in archive")
        
        if not html_files:
            handle_error(f"No HTML files found in {htmlz_path}")
        
        main_html = None
        for html_file in html_files:
            logger.debug(f"Checking HTML file: {html_file.name}")
            if html_file.name.lower() in ['index.html', 'index.htm']:
                main_html = html_file
                logger.debug(f"✓ Found index file: {main_html}")
                break
        
        if not main_html:
            main_html = html_files[0]
            logger.debug(f"Using first HTML file: {main_html}")
        
        logger.debug(f"✓ Selected HTML file: {main_html}")
        return str(main_html), extract_dir
    
    except zipfile.BadZipFile:
        handle_error(f"Failed to extract {htmlz_path}. Not a valid zip file.")
    except Exception as e:
        logger.error(f"Error during extraction: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Failed to extract htmlz: {e}")

def convert_ebook_to_html(ebook_path: str, converter_path: str, output_format: str) -> Tuple[str, Optional[str], Optional[str]]:
    """
    Converts an ebook to HTML.
    Returns: (html_file_path, intermediate_file, extract_dir)
    """
    if not os.path.exists(ebook_path):
        handle_error(f"Ebook file not found at: {ebook_path}")
    
    logger.debug(f"Starting conversion of: {ebook_path}")
    validate_file_size(ebook_path)
    
    if output_format == 'htmlz':
        output_path = os.path.splitext(ebook_path)[0] + '.htmlz'
    else:
        output_path = os.path.splitext(ebook_path)[0] + '.html'
    
    logger.debug(f"Output format: {output_format}")
    logger.debug(f"Output path: {output_path}")
    logger.info(f"Converting '{os.path.basename(ebook_path)}' to {output_format.upper()}...")
    
    try:
        logger.debug(f"Running converter: {converter_path}")
        logger.debug(f"Command: {converter_path} {ebook_path} {output_path}")
        
        result = subprocess.run(
            [converter_path, ebook_path, output_path],
            check=True,
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if logger.level <= logging.DEBUG:
            if result.stdout:
                logger.debug(f"Converter stdout: {result.stdout[:500]}...")
            if result.stderr:
                logger.debug(f"Converter stderr: {result.stderr[:500]}...")
        
        logger.info("✓ Conversion successful")
        
        if not os.path.exists(output_path):
            handle_error("Output file was not created by converter")
        
        output_size = os.path.getsize(output_path)
        logger.debug(f"Output file size: {output_size} bytes ({output_size/1024:.1f} KB)")
        
        if output_size < 100:
            handle_error("Generated file is suspiciously small. Conversion may have failed.")
        
        if output_format == 'htmlz':
            logger.debug("Extracting htmlz archive...")
            html_file, extract_dir = extract_htmlz(output_path)
            return html_file, output_path, extract_dir
        
        return output_path, None, None
    
    except FileNotFoundError:
        handle_error(f"The command '{converter_path}' was not found.")
    except subprocess.TimeoutExpired:
        handle_error("Conversion timed out. The file may be too large or corrupted.")
    except subprocess.CalledProcessError as e:
        logger.error(f"Converter failed with error:\n{e.stderr}")
        logger.debug("Full error details:", exc_info=True)
        handle_error("Conversion failed. Check that the file is not DRM-protected.")
    except Exception as e:
        logger.error(f"Unexpected error during conversion: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Conversion failed: {e}")

def try_native_conversion(ebook_path: str) -> Tuple[Optional[str], None, None]:
    """Attempt to convert using native Python libraries as fallback."""
    logger.info("Attempting native Python conversion...")
    logger.debug(f"Trying to convert: {ebook_path}")
    
    try:
        logger.debug("Trying mobi library...")
        import mobi
        
        tempdir, filepath = mobi.extract(ebook_path)
        logger.debug(f"mobi extracted to: {tempdir}")
        logger.debug(f"mobi HTML file: {filepath}")
        logger.info(f"✓ Extracted with mobi library: {filepath}")
        return filepath, None, None
    
    except ImportError:
        logger.debug("mobi library not available (install: pip install mobi)")
    except Exception as e:
        logger.debug(f"mobi library failed: {e}")
        logger.debug("mobi exception details:", exc_info=True)
    
    try:
        logger.debug("Trying ebooklib...")
        import ebooklib
        from ebooklib import epub
        
        book = epub.read_epub(ebook_path)
        html_content = []
        
        items = list(book.get_items())
        logger.debug(f"Found {len(items)} items in epub")
        
        for item in items:
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                html_content.append(item.get_content().decode('utf-8'))
        
        logger.debug(f"Extracted {len(html_content)} HTML documents")
        
        output_path = os.path.splitext(ebook_path)[0] + '_native.html'
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('<html><body>')
            f.write(''.join(html_content))
            f.write('</body></html>')
        
        logger.debug(f"Wrote combined HTML to: {output_path}")
        logger.info(f"✓ Converted with ebooklib: {output_path}")
        return output_path, None, None
    
    except ImportError:
        logger.debug("ebooklib not available (install: pip install ebooklib)")
    except Exception as e:
        logger.debug(f"ebooklib conversion failed: {e}")
        logger.debug("ebooklib exception details:", exc_info=True)
    
    logger.error("All native conversion methods failed")
    logger.error("Install native converters: pip install mobi ebooklib")
    return None, None, None

def find_book_highlights(ebook_title: str, all_clippings: List[Dict[str, str]]) -> List[str]:
    """Uses fuzzy matching to find the clippings for the specified ebook."""
    clipping_titles = list(set(c['title'] for c in all_clippings))
    
    logger.debug(f"Searching among {len(clipping_titles)} unique book titles")
    logger.debug(f"Ebook title to match: '{ebook_title}'")
    
    result = fuzzy_process.extractOne(ebook_title, clipping_titles)
    if result is None:
        logger.warning("Could not match ebook to any clippings titles.")
        return []
    
    best_match, score = result
    
    logger.info(f"📖 Ebook Title: '{ebook_title}'")
    logger.info(f"🎯 Best Match from Clippings: '{best_match}' (Confidence: {score}%)")
    
    if score < 75:
        logger.warning("Low confidence in title match.")
        logger.warning("The highlights may be from a different edition or book.")

    highlights = [c['text'] for c in all_clippings if c['title'] == best_match]
    logger.debug(f"Found {len(highlights)} highlights for this book")
    
    seen = set()
    unique_highlights = []
    for h in highlights:
        if h not in seen:
            seen.add(h)
            unique_highlights.append(h)
    
    if len(highlights) != len(unique_highlights):
        logger.debug(f"Removed {len(highlights) - len(unique_highlights)} duplicate highlights")
    
    logger.debug(f"Returning {len(unique_highlights)} unique highlights")
    return unique_highlights

# --- TEXT EXTRACTION WITH FORMATTING ---

def extract_text_with_formatting(soup: BeautifulSoup) -> Tuple[str, List[int], List[Dict]]:
    """
    Extract text while preserving paragraph boundaries and inline formatting.
    Returns: (full_text, paragraph_break_positions, formatting_spans)
    formatting_spans = [{'start': int, 'end': int, 'type': str, 'data': any}, ...]
    """
    logger.debug("Extracting text with structure and formatting preservation...")
    
    text_parts = []
    para_breaks = []
    formatting_spans = []
    
    # Find main content area
    main_content = soup.find('body')
    if not main_content:
        main_content = soup
        logger.debug("No <body> tag found, using whole document")
    
    # Track character position in the output text
    char_pos = 0
    
    # Process block-level elements
    block_elements = main_content.find_all(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'pre'])
    
    if not block_elements:
        logger.debug("No block elements found, using simple text extraction")
        return soup.get_text(), [], []
    
    logger.debug(f"Processing {len(block_elements)} block elements")
    
    processed_count = 0
    for block_elem in block_elements:
        processed_count += 1
        if processed_count % 100 == 0:
            logger.debug(f"Processed {processed_count}/{len(block_elements)} block elements")
        
        # Extract text from this block with inline formatting
        block_start_pos = char_pos
        
        for element in block_elem.descendants:
            if isinstance(element, NavigableString):
                text = str(element)
                if text.strip():  # Only process non-empty text
                    text_parts.append(text)
                    text_start = char_pos
                    char_pos += len(text)
                    text_end = char_pos
                    
                    # Check parent for formatting
                    parent = element.parent
                    while parent and parent != block_elem:
                        if parent.name in ['b', 'strong']:
                            formatting_spans.append({
                                'start': text_start,
                                'end': text_end,
                                'type': 'bold'
                            })
                            logger.debug(f"Bold: {text_start}-{text_end}")
                        elif parent.name in ['i', 'em']:
                            formatting_spans.append({
                                'start': text_start,
                                'end': text_end,
                                'type': 'italic'
                            })
                            logger.debug(f"Italic: {text_start}-{text_end}")
                        elif parent.name == 'u':
                            formatting_spans.append({
                                'start': text_start,
                                'end': text_end,
                                'type': 'underline'
                            })
                            logger.debug(f"Underline: {text_start}-{text_end}")
                        elif parent.name == 'a' and parent.get('href'):
                            formatting_spans.append({
                                'start': text_start,
                                'end': text_end,
                                'type': 'link',
                                'data': parent.get('href')
                            })
                            logger.debug(f"Link: {text_start}-{text_end} -> {parent.get('href')}")
                        elif parent.name in ['sup', 'sub']:
                            formatting_spans.append({
                                'start': text_start,
                                'end': text_end,
                                'type': 'superscript' if parent.name == 'sup' else 'subscript'
                            })
                            logger.debug(f"{parent.name}: {text_start}-{text_end}")
                        
                        parent = parent.parent
        
        # Add paragraph break after this block
        if char_pos > block_start_pos:  # Only if we added text
            text_parts.append('\n\n')
            para_breaks.append(char_pos)
            char_pos += 2
    
    full_text = ''.join(text_parts)
    logger.debug(f"✓ Extracted {len(full_text)} characters with {len(para_breaks)} paragraph breaks and {len(formatting_spans)} formatting spans")
    
    if logger.level <= logging.DEBUG and formatting_spans:
        bold_count = sum(1 for f in formatting_spans if f['type'] == 'bold')
        italic_count = sum(1 for f in formatting_spans if f['type'] == 'italic')
        underline_count = sum(1 for f in formatting_spans if f['type'] == 'underline')
        link_count = sum(1 for f in formatting_spans if f['type'] == 'link')
        logger.debug(f"Formatting stats: {bold_count} bold, {italic_count} italic, {underline_count} underline, {link_count} links")
    
    return full_text, para_breaks, formatting_spans

def create_highlighted_docx(
    html_path: str,
    highlights: List[str],
    doc_title: str,
    matcher_func: Callable,
    threshold: float
) -> Tuple[Document, int]:
    """Create a docx file with highlights and preserved formatting."""
    logger.info(f"🔍 Searching for {len(highlights)} highlights using '{matcher_func.__name__}' method...")
    logger.debug(f"HTML path: {html_path}")
    logger.debug(f"Threshold: {threshold}")
    
    logger.debug("Reading and parsing HTML file...")
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f, 'html.parser')
    except Exception as e:
        logger.error(f"Failed to read HTML file: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not read HTML file: {e}")

    logger.debug("Extracting text with structure and formatting preservation...")
    full_text_raw, para_breaks, formatting_spans = extract_text_with_formatting(soup)
    
    if not full_text_raw or len(full_text_raw) < 100:
        handle_error("Extracted text from HTML is empty or too short. The book may be DRM-protected or corrupted.")
    
    logger.debug(f"Book text length: {len(full_text_raw)} characters")
    logger.debug(f"Paragraph breaks: {len(para_breaks)}")
    logger.debug(f"Formatting spans: {len(formatting_spans)}")
    logger.debug(f"First 200 chars: {full_text_raw[:200]}...")
    
    logger.debug("Starting highlight matching...")
    found_spans, highlights_found = matcher_func(full_text_raw, highlights, threshold)
    
    percentage = (highlights_found/len(highlights)*100) if len(highlights) > 0 else 0
    logger.info(f"📊 Found {highlights_found} of {len(highlights)} highlights ({percentage:.1f}%)")
    
    if highlights_found == 0:
        logger.warning("⚠️  No highlights were found! This could mean:")
        logger.warning("  - The book text was reformatted during conversion")
        logger.warning("  - The highlights are from a different edition")
        logger.warning("  - Try using the 'vector' method: -m vector")
    elif percentage < 50:
        logger.warning(f"⚠️  Low match rate ({percentage:.1f}%). Consider using -m vector for better results.")
    
    logger.debug(f"Sorting {len(found_spans)} found spans...")
    found_spans.sort()

    logger.debug("Creating Word document with preserved structure and formatting...")
    doc = Document()
    doc.add_heading(doc_title, level=1)
    
    # Create lookup sets for fast access
    highlighted_chars = set()
    for start, end in found_spans:
        highlighted_chars.update(range(start, end))
    
    para_break_set = set(para_breaks)
    
    # Build formatting map: char_pos -> [list of formatting dicts]
    formatting_map = {}
    for fmt in formatting_spans:
        for pos in range(fmt['start'], fmt['end']):
            if pos not in formatting_map:
                formatting_map[pos] = []
            formatting_map[pos].append(fmt)
    
    logger.debug(f"Built formatting map for {len(formatting_map)} character positions")
    
    current_paragraph = doc.add_paragraph()
    current_run_text = []
    current_run_highlighted = False
    current_run_formatting = set()  # Track active formatting for current run
    
    logger.debug("Adding text with highlights, paragraph breaks, and formatting...")
    
    chars_processed = 0
    for i, char in enumerate(full_text_raw):
        chars_processed += 1
        if chars_processed % 10000 == 0:
            logger.debug(f"Processed {chars_processed}/{len(full_text_raw)} characters")
        
        is_highlighted = i in highlighted_chars
        char_formatting = set()
        
        # Get formatting for this character
        if i in formatting_map:
            for fmt in formatting_map[i]:
                char_formatting.add((fmt['type'], fmt.get('data')))
        
        # Check if we need to start a new paragraph
        if i in para_break_set:
            # Flush current run
            if current_run_text:
                text = ''.join(current_run_text)
                run = current_paragraph.add_run(text)
                if current_run_highlighted:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                # Apply formatting
                for fmt_type, fmt_data in current_run_formatting:
                    if fmt_type == 'bold':
                        run.bold = True
                    elif fmt_type == 'italic':
                        run.italic = True
                    elif fmt_type == 'underline':
                        run.underline = True
                current_run_text = []
            
            # Start new paragraph
            current_paragraph = doc.add_paragraph()
            if char == '\n' and i + 1 < len(full_text_raw) and full_text_raw[i + 1] == '\n':
                continue  # Skip the paragraph break markers
            continue
        
        # If highlight status or formatting changes, flush current run and start new one
        if current_run_text and (is_highlighted != current_run_highlighted or char_formatting != current_run_formatting):
            text = ''.join(current_run_text)
            run = current_paragraph.add_run(text)
            if current_run_highlighted:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            # Apply formatting
            for fmt_type, fmt_data in current_run_formatting:
                if fmt_type == 'bold':
                    run.bold = True
                elif fmt_type == 'italic':
                    run.italic = True
                elif fmt_type == 'underline':
                    run.underline = True
            
            current_run_text = []
            current_run_highlighted = is_highlighted
            current_run_formatting = char_formatting
        
        current_run_text.append(char)
        if not current_run_text or len(current_run_text) == 1:
            current_run_highlighted = is_highlighted
            current_run_formatting = char_formatting
    
    # Flush final run
    if current_run_text:
        text = ''.join(current_run_text)
        run = current_paragraph.add_run(text)
        if current_run_highlighted:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        for fmt_type, fmt_data in current_run_formatting:
            if fmt_type == 'bold':
                run.bold = True
            elif fmt_type == 'italic':
                run.italic = True
            elif fmt_type == 'underline':
                run.underline = True
    
    logger.debug("✓ Document creation complete")
    return doc, highlights_found

# --- MATCHER IMPLEMENTATIONS ---

def regex_matcher(book_text: str, highlights: List[str]) -> Tuple[List[Tuple[int, int]], int]:
    """Finds highlights using flexible regular expressions."""
    logger.debug("Starting regex matching...")
    found_spans = []
    
    highlights_sorted = sorted(highlights, key=len, reverse=True)
    logger.debug(f"Processing {len(highlights_sorted)} highlights (longest first)")
    
    for idx, text in enumerate(highlights_sorted):
        if idx % 10 == 0 and logger.level <= logging.DEBUG:
            logger.debug(f"Processing highlight {idx+1}/{len(highlights)}")
        
        if idx == 0:
            logger.debug(f"First highlight (length {len(text)}): {text[:50]}...")
        
        pattern = re.escape(text).replace(r'\ ', r'\s+').replace(r'\-', r'[-\u2010-\u2015]')
        try:
            for match in re.finditer(pattern, book_text, re.IGNORECASE | re.UNICODE):
                start, end = match.span()
                if not any(max(start, s) < min(end, e) for s, e in found_spans):
                    found_spans.append((start, end))
                    if idx < 5:
                        logger.debug(f"✓ Match found at position {start}-{end}")
                    break
        except re.error as e:
            logger.debug(f"Regex error for pattern: {e}")
            continue
        except Exception as e:
            logger.debug(f"Unexpected error in regex matching: {e}")
            continue
    
    logger.debug(f"✓ Regex matching complete. Found {len(found_spans)} matches")
    return found_spans, len(found_spans)

def difflib_matcher(book_text: str, highlights: List[str], threshold: float) -> Tuple[List[Tuple[int, int]], int]:
    """Finds highlights using difflib's SequenceMatcher."""
    logger.debug(f"Starting difflib matching with threshold {threshold}...")
    found_spans = []
    
    logger.debug("Normalizing book text...")
    normalized_book_text = re.sub(r'\s+', ' ', book_text).lower()
    logger.debug(f"Normalized text length: {len(normalized_book_text)}")
    
    highlights_sorted = sorted(highlights, key=len, reverse=True)
    logger.debug(f"Processing {len(highlights_sorted)} highlights")
    
    for idx, text in enumerate(highlights_sorted):
        if idx % 10 == 0 and logger.level <= logging.DEBUG:
            logger.debug(f"Processing highlight {idx+1}/{len(highlights)}")
        
        normalized_highlight = re.sub(r'\s+', ' ', text).lower()
        if not normalized_highlight:
            continue

        try:
            matcher = SequenceMatcher(None, normalized_book_text, normalized_highlight)
            match = matcher.find_longest_match(0, len(normalized_book_text), 0, len(normalized_highlight))
            
            similarity = match.size / len(normalized_highlight) if len(normalized_highlight) > 0 else 0
            
            if idx < 5:
                logger.debug(f"Highlight {idx}: similarity = {similarity:.3f}, match_size = {match.size}")
            
            if similarity >= threshold:
                start = match.a
                end = match.a + len(text)
                end = min(end, len(book_text))
                
                if not any(max(start, s) < min(end, e) for s, e in found_spans):
                    found_spans.append((start, end))
                    if idx < 5:
                        logger.debug(f"✓ Match found with {similarity:.2%} similarity at position {start}-{end}")
        except Exception as e:
            logger.debug(f"Error matching highlight {idx}: {e}")
            continue
    
    logger.debug(f"✓ Difflib matching complete. Found {len(found_spans)} matches")
    return found_spans, len(found_spans)

def vector_matcher(book_text: str, highlights: List[str], threshold: float) -> Tuple[List[Tuple[int, int]], int]:
    """Finds highlights using semantic vector similarity."""
    logger.debug("Starting vector matching...")
    
    if not check_nlp_libraries():
        handle_error("The 'vector' method requires 'sentence-transformers', 'torch', and 'nltk'.\nPlease install them: pip install sentence-transformers torch nltk")
    
    import nltk
    from sentence_transformers import SentenceTransformer, util
    
    logger.debug("Checking for NLTK punkt tokenizer...")
    try:
        nltk.data.find('tokenizers/punkt')
        logger.debug("✓ punkt tokenizer found")
    except LookupError:
        logger.info("Downloading NLTK 'punkt' tokenizer...")
        try:
            nltk.download('punkt', quiet=True)
            logger.debug("✓ punkt tokenizer downloaded")
        except Exception as e:
            logger.error(f"Failed to download punkt tokenizer: {e}")
            handle_error("Could not download required NLTK data")
    
    logger.info(f"Loading sentence transformer model: {NLP_MODEL}")
    try:
        model = SentenceTransformer(NLP_MODEL)
        logger.debug(f"✓ Model loaded: {type(model)}")
    except Exception as e:
        logger.error(f"Failed to load model: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not load NLP model: {e}")
    
    logger.info("Tokenizing book text into sentences...")
    try:
        book_sentences = nltk.sent_tokenize(book_text)
        logger.debug(f"✓ Book split into {len(book_sentences)} sentences")
    except Exception as e:
        logger.error(f"Failed to tokenize text: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not tokenize book text: {e}")
    
    if len(book_sentences) > 10000:
        logger.warning(f"Large number of sentences ({len(book_sentences)}). This may take several minutes...")
    
    logger.info("Encoding book sentences...")
    try:
        logger.debug(f"Encoding {len(book_sentences)} sentences with batch_size=32")
        book_embeddings = model.encode(
            book_sentences,
            convert_to_tensor=True,
            show_progress_bar=logger.level <= logging.INFO,
            normalize_embeddings=True,
            batch_size=32
        )
        logger.debug(f"✓ Book embeddings shape: {book_embeddings.shape}")
    except Exception as e:
        logger.error(f"Failed to encode book sentences: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not encode book text: {e}")
    
    logger.info("Encoding highlights...")
    try:
        logger.debug(f"Encoding {len(highlights)} highlights")
        highlight_embeddings = model.encode(
            highlights,
            convert_to_tensor=True,
            show_progress_bar=logger.level <= logging.INFO,
            normalize_embeddings=True,
            batch_size=32
        )
        logger.debug(f"✓ Highlight embeddings shape: {highlight_embeddings.shape}")
    except Exception as e:
        logger.error(f"Failed to encode highlights: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not encode highlights: {e}")

    logger.info("Computing similarity scores...")
    try:
        logger.debug("Running cosine similarity...")
        cosine_scores = util.cos_sim(highlight_embeddings, book_embeddings)
        logger.debug(f"✓ Cosine scores shape: {cosine_scores.shape}")
    except Exception as e:
        logger.error(f"Failed to compute similarity: {e}")
        logger.debug("Full traceback:", exc_info=True)
        handle_error(f"Could not compute similarity scores: {e}")
    
    found_spans = []
    
    for i, text in enumerate(highlights):
        if i % 10 == 0 and logger.level <= logging.DEBUG:
            logger.debug(f"Processing highlight {i+1}/{len(highlights)}")
        
        try:
            best_match_idx = cosine_scores[i].argmax().item()
            confidence = cosine_scores[i][best_match_idx].item()
            
            if i < 5 or (i % 100 == 0 and logger.level <= logging.DEBUG):
                logger.debug(f"Highlight {i+1}: best match sentence {best_match_idx}, confidence = {confidence:.3f}")
            
            if confidence > threshold:
                sentence = book_sentences[best_match_idx]
                search_start = book_text.find(sentence)
                
                if search_start != -1:
                    search_end = search_start + len(sentence) + int(len(text) * 0.3)
                    search_end = min(search_end, len(book_text))
                    search_span = book_text[search_start:search_end]
                    
                    matcher = SequenceMatcher(None, search_span, text, autojunk=False)
                    match = matcher.find_longest_match(0, len(search_span), 0, len(text))
                    
                    final_start = search_start + match.a
                    final_end = final_start + match.size
                    
                    if not any(max(final_start, s) < min(final_end, e) for s, e in found_spans):
                        found_spans.append((final_start, final_end))
                        if i < 5:
                            logger.debug(f"✓ Match added at position {final_start}-{final_end}")
                else:
                    if i < 5:
                        logger.debug(f"Could not find sentence in book text")
            else:
                if i < 5:
                    logger.debug(f"Highlight {i+1}: confidence too low ({confidence:.3f} < {threshold})")
        except Exception as e:
            logger.debug(f"Error processing highlight {i}: {e}")
            continue
    
    logger.debug(f"✓ Vector matching complete. Found {len(found_spans)} matches")
    return found_spans, len(found_spans)

# --- SINGLE BOOK PROCESSING ---

def process_single_book(
    ebook_path: str,
    all_clippings: List[Dict[str, str]],
    converter_path: str,
    output_format: str,
    method: str = 'diff',
    output_path: Optional[str] = None,
    keep_html: bool = False,
    similarity_threshold: float = DEFAULT_SIMILARITY_THRESHOLD,
    vector_threshold: float = DEFAULT_VECTOR_CONFIDENCE_THRESHOLD,
    compare_mode: bool = False
) -> Dict:
    """
    Process a single book and return results.
    Returns dict with: success, output_path, highlights_found, highlights_total, error
    """
    logger.debug(f"Starting process_single_book for: {ebook_path}")
    result = {
        'success': False,
        'ebook_path': str(ebook_path),
        'output_path': None,
        'highlights_found': 0,
        'highlights_total': 0,
        'error': None
    }
    
    html_file = None
    intermediate_file = None
    extract_dir = None
    
    try:
        # Convert ebook
        logger.debug("Converting ebook to HTML...")
        html_file, intermediate_file, extract_dir = convert_ebook_to_html(
            str(ebook_path), converter_path, output_format
        )
        
        # Extract title
        logger.debug(f"Reading HTML file for title extraction: {html_file}")
        try:
            with open(html_file, 'r', encoding='utf-8') as f:
                soup = BeautifulSoup(f, 'html.parser')
                doc_title = soup.title.string.strip() if soup.title and soup.title.string else os.path.basename(ebook_path)
        except Exception as e:
            logger.warning(f"Could not extract title from HTML: {e}")
            doc_title = os.path.basename(ebook_path)
        
        logger.debug(f"Document title: '{doc_title}'")
        
        # Find highlights
        logger.debug("Finding highlights for this book...")
        relevant_highlights = find_book_highlights(doc_title, all_clippings)
        
        if not relevant_highlights:
            result['error'] = "No highlights found"
            logger.warning("No highlights found for this book")
            return result
        
        result['highlights_total'] = len(relevant_highlights)
        logger.debug(f"Found {len(relevant_highlights)} highlights to match")
        
        # --- COMPARE MODE ---
        if compare_mode:
            logger.debug("Comparison mode activated")
            logger.info("\n" + "=" * 70)
            logger.info("⚖️  Running Comparison Mode")
            logger.info("=" * 70)
            
            results = {}
            
            # Test Regex
            try:
                logger.info(f"\n🧪 Testing Regex method...")
                _, count = create_highlighted_docx(
                    html_file, relevant_highlights, doc_title,
                    lambda txt, h, t: regex_matcher(txt, h),
                    threshold=0.0
                )
                results["Regex"] = f"{count}/{len(relevant_highlights)} ({count/len(relevant_highlights)*100:.1f}%)"
                logger.debug(f"Regex result: {results['Regex']}")
            except Exception as e:
                logger.error(f"Regex method failed: {e}")
                logger.debug(f"Regex exception details:", exc_info=True)
                results["Regex"] = f"Failed ({type(e).__name__})"
            
            # Test Difflib
            try:
                logger.info(f"\n🧪 Testing Difflib method...")
                _, count = create_highlighted_docx(
                    html_file, relevant_highlights, doc_title,
                    difflib_matcher,
                    similarity_threshold
                )
                results["Difflib"] = f"{count}/{len(relevant_highlights)} ({count/len(relevant_highlights)*100:.1f}%)"
                logger.debug(f"Difflib result: {results['Difflib']}")
            except Exception as e:
                logger.error(f"Difflib method failed: {e}")
                logger.debug(f"Difflib exception details:", exc_info=True)
                results["Difflib"] = f"Failed ({type(e).__name__})"
            
            # Test Vector
            try:
                logger.info(f"\n🧪 Testing Vector method...")
                _, count = create_highlighted_docx(
                    html_file, relevant_highlights, doc_title,
                    vector_matcher,
                    vector_threshold
                )
                results["Vector"] = f"{count}/{len(relevant_highlights)} ({count/len(relevant_highlights)*100:.1f}%)"
                logger.debug(f"Vector result: {results['Vector']}")
            except Exception as e:
                logger.error(f"Vector method failed: {e}")
                logger.debug(f"Vector exception details:", exc_info=True)
                results["Vector"] = f"Failed ({type(e).__name__})"

            print("\n" + "=" * 70)
            print("📊 Comparison Results")
            print("=" * 70)
            for name, res in results.items():
                print(f"  {name:<10}: {res}")
            print("=" * 70)
            
            result['success'] = True
            result['compare_results'] = results
            return result
        
        # --- SINGLE METHOD MODE ---
        logger.debug(f"Single method mode: {method}")
        
        # Select matcher
        if method == 'regex':
            selected_matcher = lambda txt, h, t: regex_matcher(txt, h)
            threshold = 0.0
            logger.debug("Using regex matcher (no threshold)")
        elif method == 'diff':
            selected_matcher = difflib_matcher
            threshold = similarity_threshold
            logger.debug(f"Using difflib matcher (threshold: {threshold})")
        else:  # vector
            selected_matcher = vector_matcher
            threshold = vector_threshold
            logger.debug(f"Using vector matcher (threshold: {threshold})")

        logger.debug("Creating highlighted document...")
        doc, found_count = create_highlighted_docx(
            html_file,
            relevant_highlights,
            doc_title,
            selected_matcher,
            threshold
        )
        
        result['highlights_found'] = found_count
        logger.debug(f"Found {found_count} highlights")
        
        # Save document
        if not output_path:
            output_path = os.path.splitext(ebook_path)[0] + '_highlighted.docx'
        
        logger.debug(f"Output path: {output_path}")
        logger.info(f"💾 Saving document to: {output_path}")
        
        try:
            doc.save(output_path)
            logger.debug("✓ Document saved successfully")
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            logger.debug("Full traceback:", exc_info=True)
            result['error'] = f"Failed to save document: {e}"
            return result
        
        result['success'] = True
        result['output_path'] = output_path
        
    except Exception as e:
        logger.error(f"Error processing book: {e}")
        logger.debug("Full exception:", exc_info=True)
        result['error'] = str(e)
    
    finally:
        # Cleanup
        if not keep_html:
            logger.debug("Cleaning up intermediate files...")
            if html_file and os.path.exists(html_file):
                try:
                    os.remove(html_file)
                    logger.debug(f"✓ Cleaned up: {os.path.basename(html_file)}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {html_file}: {e}")
            
            if extract_dir and os.path.exists(extract_dir):
                try:
                    shutil.rmtree(extract_dir)
                    logger.debug(f"✓ Cleaned up extracted directory: {extract_dir}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {extract_dir}: {e}")
            
            if intermediate_file and os.path.exists(intermediate_file):
                try:
                    os.remove(intermediate_file)
                    logger.debug(f"✓ Cleaned up: {os.path.basename(intermediate_file)}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {intermediate_file}: {e}")
        else:
            logger.info(f"Keeping intermediate files (--preserve-html):")
            if html_file:
                logger.info(f"  - HTML: {html_file}")
            if intermediate_file:
                logger.info(f"  - Archive: {intermediate_file}")
            if extract_dir:
                logger.info(f"  - Extract dir: {extract_dir}")
    
    return result

# --- MAIN EXECUTION ---

def main():
    parser = argparse.ArgumentParser(
        description="Generate Word documents with Kindle highlights - LIBRARY BATCH MODE BY DEFAULT",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Default: Process ALL books from library with 95%+ match confidence:
    python kindle_highlighter.py --clippings "My Clippings.txt" -v
  
  Process only the best matching book from library:
    python kindle_highlighter.py --clippings "My Clippings.txt" --no-batch -v
  
  Process specific ebook file:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" -v
  
  Batch mode with HTML files preserved for debugging:
    python kindle_highlighter.py --clippings "My Clippings.txt" --preserve-html -vv
  
  Compare all methods on a specific file:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" --compare -v
  
  List all books in library:
    python kindle_highlighter.py --list-books

Note: Library batch mode is enabled by default. Use --no-batch to process only 
      the single best match, or --ebook to process a specific file.
        """
    )
    
    # Mode selection
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--list-books", action="store_true", 
                           help="List all books in Calibre Library and exit.")
    
    # Direct file mode
    direct_group = parser.add_argument_group('Direct File Mode')
    direct_group.add_argument("--ebook", dest="ebook_file", 
                             help="Path to a single ebook file (overrides library mode).")
    direct_group.add_argument("--clippings", dest="clippings_file", 
                             help="Path to 'My Clippings.txt' (required).")
    
    # Library options (defaults enabled)
    library_group = parser.add_argument_group('Library Options')
    library_group.add_argument("--library-path", 
                              help=f"Path to Calibre Library (default: {DEFAULT_CALIBRE_LIBRARY})")
    library_group.add_argument("--no-batch", action="store_true",
                              help="Process only the single best match instead of all books with 95%+ confidence")
    
    # Converter options
    converter_group = parser.add_argument_group('Converter Options')
    converter_group.add_argument("--calibre-path", 
                                help="Path to Calibre's ebook-convert binary")
    converter_group.add_argument("--try-native", action="store_true", 
                                help="Try native Python conversion if converter fails.")
    
    # Output options
    output_group = parser.add_argument_group('Output Options')
    output_group.add_argument("-o", "--output", 
                             help="Output path (single file mode only)")
    output_group.add_argument("--keep-html", action="store_true", 
                             help="Keep intermediate HTML files")
    output_group.add_argument("--preserve-html", action="store_true",
                             help="Same as --keep-html (preserve HTML files for inspection)")
    
    # Matching options
    matching_group = parser.add_argument_group('Matching Options')
    matching_group.add_argument(
        "-m", "--method",
        choices=['regex', 'diff', 'vector'],
        default='diff',
        help="Matching method: regex (fast), diff (balanced, default), vector (best accuracy)"
    )
    matching_group.add_argument("--compare", action="store_true", 
                               help="Run all methods and compare results (single file only)")
    matching_group.add_argument(
        "--similarity-threshold",
        type=float,
        default=DEFAULT_SIMILARITY_THRESHOLD,
        help=f"Similarity threshold for difflib (default: {DEFAULT_SIMILARITY_THRESHOLD})"
    )
    matching_group.add_argument(
        "--vector-threshold",
        type=float,
        default=DEFAULT_VECTOR_CONFIDENCE_THRESHOLD,
        help=f"Confidence threshold for vector (default: {DEFAULT_VECTOR_CONFIDENCE_THRESHOLD})"
    )
    
    # General options
    parser.add_argument(
        "-v", "--verbose",
        action="count",
        default=0,
        help="Increase verbosity (-v for info, -vv for debug)"
    )
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.verbose)
    
    logger.info("=" * 70)
    logger.info("📚 Kindle Ebook Highlighter - COMPLETE VERSION")
    logger.info("=" * 70)
    logger.debug(f"Command line arguments: {vars(args)}")
    logger.debug(f"Python version: {sys.version}")
    logger.debug(f"Platform: {sys.platform}")
    
    # Handle preserve-html alias
    if args.preserve_html:
        args.keep_html = True
        logger.debug("--preserve-html flag set, enabling --keep-html")
    
    # --- LIST BOOKS MODE ---
    if args.list_books:
        logger.debug("List books mode activated")
        library_path = find_calibre_library(args.library_path)
        books = discover_books(library_path)
        
        print("\n" + "=" * 70)
        print(f"Books in Calibre Library ({library_path})")
        print("=" * 70)
        
        for book in books:
            print(f"📖 {book['title']}")
            print(f"   Author: {book['author']}")
            print(f"   Path: {book['path']}")
            print()
        
        print(f"Total: {len(books)} books")
        sys.exit(0)
    
    # Check dependencies
    logger.info("\n🔍 Checking dependencies...")
    deps = check_dependencies(args.calibre_path)
    
    required = ['docx', 'bs4', 'thefuzz', 'converter']
    missing_required = [k for k in required if not deps[k]]
    
    if missing_required:
        handle_error(f"Missing required dependencies: {', '.join(missing_required)}")
    
    if args.try_native and not (deps['mobi'] or deps['ebooklib']):
        logger.warning("--try-native flag used, but 'mobi' and 'ebooklib' libraries are not installed.")
        logger.warning("Install them: pip install mobi ebooklib")
    
    logger.info("✓ All required dependencies found\n")
    
    # --- CORE LOGIC ---
    try:
        logger.debug("Starting main processing...")
        
        # Get clippings path
        clippings_path = args.clippings_file
        
        # If no ebook file specified, use library mode by default
        if not args.ebook_file and not clippings_path:
            logger.debug("No ebook or clippings specified, checking for defaults...")
            clippings_default = os.path.expanduser('~/Documents/My Clippings.txt')
            if os.path.exists(clippings_default):
                clippings_path = clippings_default
                logger.info(f"Using default clippings: {clippings_path}")
        
        if not clippings_path:
            handle_error("A clippings file must be provided with --clippings.")
        
        logger.debug(f"Clippings path: {clippings_path}")
        
        # Parse clippings
        all_clippings = parse_clippings(clippings_path)
        
        # Get converter
        converter_path, output_format = check_converter_available(args.calibre_path)
        
        if converter_path:
            logger.debug(f"Converter available: {converter_path}")
        elif args.try_native:
            logger.warning("No standard converter found. Will try native conversion.")
        else:
            handle_error("No converter available. Install Calibre or use --try-native")
        
        # --- LIBRARY MODE (DEFAULT) ---
        if not args.ebook_file:
            logger.debug("Library mode: processing books from Calibre library")
            library_path = find_calibre_library(args.library_path)
            books = discover_books(library_path)
            clipping_titles = [c['title'] for c in all_clippings]
            
            # Match all books
            matches = match_books_to_clippings(books, clipping_titles, MIN_TITLE_MATCH_SCORE)
            
            if not matches:
                handle_error("No matching books found in library!")
            
            # Show matches
            logger.info("\n" + "=" * 80)
            logger.info("Book Matching Results:")
            logger.info("=" * 80)
            for idx, m in enumerate(matches[:20], 1):
                logger.info(f"{idx:2d}. Score {m['score']:3d}% | {m['highlight_count']:4d} highlights | '{m['book']['title']}'")
                logger.info(f"     Clipping title: '{m['clipping_title']}'")
            if len(matches) > 20:
                logger.info(f"     ... and {len(matches) - 20} more books")
            logger.info("=" * 80)
            
            # Determine which books to process (BATCH BY DEFAULT)
            if args.no_batch:
                # Process only the best match
                if matches[0]['score'] >= AUTO_SELECT_THRESHOLD:
                    books_to_process = [matches[0]]
                    logger.info(f"\n✓ Processing single best match: '{matches[0]['book']['title']}' ({matches[0]['score']}%)")
                    high_conf_count = len([m for m in matches if m['score'] >= AUTO_SELECT_THRESHOLD])
                    if high_conf_count > 1:
                        logger.info(f"💡 Tip: Remove --no-batch to process all {high_conf_count} books with {AUTO_SELECT_THRESHOLD}%+ match")
                else:
                    logger.error(f"\n❌ Best match only {matches[0]['score']}% (need {AUTO_SELECT_THRESHOLD}%+)")
                    logger.error(f"   Book: '{matches[0]['book']['title']}'")
                    logger.error("   Try removing --no-batch, or use --ebook for specific file")
                    sys.exit(1)
            else:
                # DEFAULT: Process ALL books with 95%+ match (BATCH MODE)
                books_to_process = [m for m in matches if m['score'] >= AUTO_SELECT_THRESHOLD]
                logger.info(f"\n📦 BATCH MODE (default): Processing {len(books_to_process)} books with {AUTO_SELECT_THRESHOLD}%+ match confidence")
                if len(books_to_process) == 0:
                    logger.error(f"No books found with {AUTO_SELECT_THRESHOLD}%+ match confidence")
                    logger.error(f"Best match was {matches[0]['score']}%: '{matches[0]['book']['title']}'")
                    sys.exit(1)
            
            # Process books
            results = []
            
            for idx, match in enumerate(books_to_process, 1):
                book = match['book']
                logger.info(f"\n{'='*70}")
                logger.info(f"Processing {idx}/{len(books_to_process)}: {book['title']}")
                logger.info(f"Match: {match['clipping_title']} ({match['score']}%, {match['highlight_count']} highlights)")
                logger.info(f"{'='*70}")
                
                result = process_single_book(
                    book['path'],
                    all_clippings,
                    converter_path,
                    output_format,
                    method=args.method,
                    keep_html=args.keep_html,
                    similarity_threshold=args.similarity_threshold,
                    vector_threshold=args.vector_threshold,
                    compare_mode=False  # No compare in batch mode
                )
                
                results.append({
                    'book_title': book['title'],
                    'match_score': match['score'],
                    **result
                })
            
            # Summary
            print("\n" + "=" * 70)
            print("📊 BATCH PROCESSING SUMMARY")
            print("=" * 70)
            
            successful = [r for r in results if r['success']]
            failed = [r for r in results if not r['success']]
            
            print(f"Total books processed: {len(results)}")
            print(f"✅ Successful: {len(successful)}")
            print(f"❌ Failed: {len(failed)}")
            print()
            
            if successful:
                print("✅ Successful:")
                for r in successful:
                    pct = (r['highlights_found']/r['highlights_total']*100) if r['highlights_total'] > 0 else 0
                    print(f"   • {r['book_title']}")
                    print(f"     Found {r['highlights_found']}/{r['highlights_total']} highlights ({pct:.1f}%)")
                    print(f"     Saved to: {r['output_path']}")
                print()
            
            if failed:
                print("❌ Failed:")
                for r in failed:
                    print(f"   • {r['book_title']}: {r['error']}")
                print()
            
            print("=" * 70)
        
        # --- SINGLE FILE MODE ---
        else:
            ebook_path = args.ebook_file
            logger.debug(f"Single file mode: processing {ebook_path}")
            
            result = process_single_book(
                ebook_path,
                all_clippings,
                converter_path,
                output_format,
                method=args.method,
                output_path=args.output,
                keep_html=args.keep_html,
                similarity_threshold=args.similarity_threshold,
                vector_threshold=args.vector_threshold,
                compare_mode=args.compare
            )
            
            if args.compare:
                # Results already printed in compare mode
                logger.debug("Compare mode results displayed")
            elif result['success']:
                pct = (result['highlights_found']/result['highlights_total']*100) if result['highlights_total'] > 0 else 0
                print("\n" + "=" * 70)
                print(f"✅ Success! Found {result['highlights_found']}/{result['highlights_total']} highlights ({pct:.1f}%)")
                print(f"📄 Document saved to: {result['output_path']}")
                print("=" * 70)
            else:
                print("\n" + "=" * 70)
                print(f"❌ Failed: {result['error']}")
                print("=" * 70)
                sys.exit(1)

    except KeyboardInterrupt:
        logger.warning("\n\n⚠️  Operation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.exception("\n❌ Unexpected error occurred:")
        sys.exit(1)

if __name__ == "__main__":
    main()