#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Kindle Ebook Highlighter - COMPLETE VERSION with Batch Processing

All features included:
- Batch processing (--batch flag to process ALL books)
- All 3 matching methods: regex, difflib, vector
- Compare mode
- Fixed html_path → html_file bug
- Improved book matching
- All original functionality preserved

Usage:
  Batch process ALL books:
    python kindle_highlighter.py --library --clippings "My Clippings.txt" --batch -v
  
  Process single best match:
    python kindle_highlighter.py --library --clippings "My Clippings.txt" -v
  
  Compare all methods:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" --compare -v
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
from typing import List, Dict, Tuple, Callable, Optional
from pathlib import Path
from difflib import SequenceMatcher
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from bs4 import BeautifulSoup
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
    else:
        level = logging.DEBUG
    
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(levelname)s: %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(level)

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
    # Remove author/series in parentheses at the end
    title_clean = re.sub(r'\s*\([^)]+\)\s*$', '', title)
    
    # Remove special unicode characters
    title_clean = title_clean.replace('¬', '').replace('', '')
    
    # Remove multiple spaces
    title_clean = re.sub(r'\s+', ' ', title_clean).strip()
    
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
                logger.debug(f"Found Calibre binary: {binary_path}")
                return binary_path
        
        common_paths = [
            '/Applications/calibre.app/Contents/MacOS/ebook-convert',
            '/Applications/Calibre.app/Contents/MacOS/ebook-convert',
            '/usr/local/bin/ebook-convert',
        ]
        logger.debug(f"Checking common macOS paths: {common_paths}")
        for path in common_paths:
            if os.path.exists(path) and os.access(path, os.X_OK):
                logger.debug(f"Found Calibre binary: {path}")
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
                logger.debug(f"Found Calibre binary: {path}")
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
                logger.debug(f"Found Calibre binary: {path}")
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
            logger.debug(f"Found ebook-convert in PATH: {path}")
            return path
        else:
            logger.debug(f"Command failed with return code: {result.returncode}")
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        logger.debug(f"PATH search failed: {e}")
    
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
            logger.debug("Found standalone ebook-converter")
            logger.info("Using standalone ebook-converter")
            return 'ebook-converter', 'htmlz'
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        logger.debug(f"Standalone ebook-converter not found: {e}")
    
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
        logger.debug("python-docx found")
    except ImportError:
        logger.error("python-docx not found. Install: pip install python-docx")
    
    try:
        import bs4
        deps['bs4'] = True
        logger.debug("beautifulsoup4 found")
    except ImportError:
        logger.error("beautifulsoup4 not found. Install: pip install beautifulsoup4")
    
    try:
        import thefuzz
        deps['thefuzz'] = True
        logger.debug("thefuzz found")
    except ImportError:
        logger.error("thefuzz not found. Install: pip install thefuzz python-Levenshtein")
    
    logger.debug("Checking for ebook converter...")
    converter, _ = check_converter_available(calibre_path)
    deps['converter'] = converter is not None
    if not deps['converter']:
        logger.error("No ebook converter found. Install Calibre or ebook-converter")
    
    logger.debug("Checking optional conversion libraries...")
    try:
        import mobi
        deps['mobi'] = True
        logger.debug("mobi library found (optional)")
    except ImportError:
        logger.debug("mobi library not found (optional: pip install mobi)")
    
    try:
        import ebooklib
        deps['ebooklib'] = True
        logger.debug("ebooklib found (optional)")
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
        logger.debug("NLP libraries found (torch, nltk, sentence_transformers)")
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
    
    logger.info(f"Using Calibre library: {path}")
    return path

def discover_books(library_path: Path, format_filter: str = 'mobi') -> List[Dict]:
    """Recursively discover all books in Calibre library."""
    logger.info(f"Scanning library for .{format_filter} files...")
    logger.debug(f"Starting recursive scan of: {library_path}")
    
    books = []
    format_filter = format_filter.lower()
    
    for root, dirs, files in os.walk(library_path):
        for file in files:
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
    
    logger.info(f"Found {len(books)} books in library")
    logger.debug(f"Book list: {[b['title'] for b in books[:5]]}{'...' if len(books) > 5 else ''}")
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
    
    matches = []
    
    for clip_title, highlight_count in title_counts.most_common():
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
                        f"'{normalize_title_for_matching(best_book['title'])}' ({best_score}%)")
    
    # Sort by: 1) score desc, 2) highlight count desc
    matches.sort(key=lambda x: (x['score'], x['highlight_count']), reverse=True)
    
    return matches

# --- FILE PARSING ---

def validate_file_size(filepath: str, max_size_mb: int = MAX_BOOK_SIZE_MB) -> bool:
    """Validate that file size is reasonable."""
    size_mb = os.path.getsize(filepath) / (1024 * 1024)
    logger.debug(f"File size: {size_mb:.2f}MB")
    if size_mb > max_size_mb:
        logger.warning(f"File is large ({size_mb:.2f}MB). This may take a while...")
    return True

def parse_clippings(clippings_file: str) -> List[Dict[str, str]]:
    """Parses the 'My Clippings.txt' file and returns a list of clippings."""
    if not os.path.exists(clippings_file):
        handle_error(f"Clippings file not found at: {clippings_file}")
    
    logger.info(f"Parsing clippings from: {clippings_file}")
    
    file_size = os.path.getsize(clippings_file) / (1024 * 1024)
    logger.debug(f"Clippings file size: {file_size:.2f}MB")
    
    if file_size > 10:
        logger.info(f"Large clippings file ({file_size:.2f}MB). This may take a moment...")
    
    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
    content = None
    
    for encoding in encodings:
        logger.debug(f"Trying encoding: {encoding}")
        try:
            with open(clippings_file, 'r', encoding=encoding) as f:
                content = f.read()
            logger.debug(f"Successfully read file with {encoding} encoding")
            break
        except UnicodeDecodeError:
            logger.debug(f"Failed to read with {encoding} encoding")
            continue
    
    if content is None:
        handle_error(f"Could not read clippings file with any known encoding")

    logger.debug(f"File content length: {len(content)} characters")
    logger.info("Parsing clippings (this may take a moment for large files)...")
    
    clippings = []
    entries = content.split('==========')
    logger.debug(f"Found {len(entries)} potential entries")
    
    highlight_patterns = [
        r'^-\s*Your\s+(Highlight|Note)',
        r'^-\s*(Ihre|Deine)\s+(Markierung|Notiz)',
        r'^-\s*Votre\s+(surlignement|note)',
        r'^-\s*Tu\s+(subrayado|nota)',
        r'^-\s*La\s+tua\s+(evidenziazione|nota)',
    ]
    
    logger.debug(f"Using {len(highlight_patterns)} language patterns for detection")
    
    for idx, entry in enumerate(entries):
        if idx % 100 == 0 and idx > 0:
            logger.debug(f"Processing entry {idx}/{len(entries)}")
        
        entry = entry.strip()
        if not entry:
            continue
        
        lines = entry.split('\n')
        if len(lines) < 3:
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
            if idx < 10:
                logger.debug(f"Skipping entry {idx}: line 2 doesn't match any pattern: {lines[1][:80]}")
            continue
        
        if idx < 10:
            logger.debug(f"Entry {idx} matched pattern: {matched_pattern}")
        
        clipping_text = None
        for i in range(2, len(lines)):
            if lines[i].strip():
                clipping_text = '\n'.join(lines[i:]).strip()
                break
        
        if not clipping_text:
            logger.debug(f"Skipping entry {idx}: no highlight text found")
            continue
        
        if len(clipping_text) >= MIN_HIGHLIGHT_LENGTH:
            clippings.append({'title': book_title, 'text': clipping_text})
            if idx < 10:
                logger.debug(f"Added highlight from '{book_title}': {clipping_text[:50]}...")
        else:
            logger.debug(f"Skipping entry {idx}: highlight too short ({len(clipping_text)} chars)")
    
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
        
        logger.debug("\nBooks found in clippings:")
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
        shutil.rmtree(extract_dir)
    
    try:
        logger.debug("Opening zip archive...")
        with zipfile.ZipFile(htmlz_path, 'r') as zip_ref:
            logger.debug(f"Extracting {len(zip_ref.namelist())} files...")
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
                logger.debug(f"Found index file: {main_html}")
                break
        
        if not main_html:
            main_html = html_files[0]
            logger.debug(f"Using first HTML file: {main_html}")
        
        logger.debug(f"Using HTML file: {main_html}")
        return str(main_html), extract_dir
    
    except zipfile.BadZipFile:
        handle_error(f"Failed to extract {htmlz_path}. Not a valid zip file.")
    except Exception as e:
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
        result = subprocess.run(
            [converter_path, ebook_path, output_path],
            check=True,
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if logger.level <= logging.DEBUG:
            logger.debug(f"Converter stdout: {result.stdout[:500]}...")
            if result.stderr:
                logger.debug(f"Converter stderr: {result.stderr[:500]}...")
        
        logger.info("✓ Conversion successful")
        
        if not os.path.exists(output_path):
            handle_error("Output file was not created by converter")
        
        output_size = os.path.getsize(output_path)
        logger.debug(f"Output file size: {output_size} bytes")
        
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
        handle_error("Conversion failed. Check that the file is not DRM-protected.")

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
    
    logger.error("All native conversion methods failed")
    logger.error("Install native converters: pip install mobi ebooklib")
    return None, None, None

def find_book_highlights(ebook_title: str, all_clippings: List[Dict[str, str]]) -> List[str]:
    """Uses fuzzy matching to find the clippings for the specified ebook."""
    clipping_titles = list(set(c['title'] for c in all_clippings))
    
    logger.debug(f"Searching among {len(clipping_titles)} unique book titles")
    logger.debug(f"Ebook title to match: {ebook_title}")
    
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

# --- MATCHER IMPLEMENTATIONS ---

def create_highlighted_docx(
    html_path: str,
    highlights: List[str],
    doc_title: str,
    matcher_func: Callable,
    threshold: float
) -> Tuple[Document, int]:
    """Generic function to create a docx file using a provided matcher."""
    logger.info(f"🔍 Searching for {len(highlights)} highlights using '{matcher_func.__name__}' method...")
    logger.debug(f"HTML path: {html_path}")
    logger.debug(f"Threshold: {threshold}")
    
    logger.debug("Reading and parsing HTML file...")
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    logger.debug("Extracting text from HTML...")
    full_text_raw = soup.get_text()
    
    if not full_text_raw or len(full_text_raw) < 100:
        handle_error("Extracted text from HTML is empty or too short. The book may be DRM-protected or corrupted.")
    
    logger.debug(f"Book text length: {len(full_text_raw)} characters")
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

    logger.debug("Creating Word document...")
    doc = Document()
    doc.add_heading(doc_title, level=1)
    current_pos = 0
    p = doc.add_paragraph()

    logger.debug("Adding text and highlights to document...")
    for idx, (start, end) in enumerate(found_spans):
        if idx % 10 == 0 and logger.level <= logging.DEBUG:
            logger.debug(f"Processing span {idx+1}/{len(found_spans)}")
        
        p.add_run(full_text_raw[current_pos:start])
        highlighted_run = p.add_run(full_text_raw[start:end])
        highlighted_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        current_pos = end

    p.add_run(full_text_raw[current_pos:])
    logger.debug("Document creation complete")
    return doc, highlights_found

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
                    logger.debug(f"Match found at position {start}-{end}")
                    break
        except re.error as e:
            logger.debug(f"Regex error for pattern: {e}")
            continue
    
    logger.debug(f"Regex matching complete. Found {len(found_spans)} matches")
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

        matcher = SequenceMatcher(None, normalized_book_text, normalized_highlight)
        match = matcher.find_longest_match(0, len(normalized_book_text), 0, len(normalized_highlight))
        
        similarity = match.size / len(normalized_highlight) if len(normalized_highlight) > 0 else 0
        
        if idx < 5:
            logger.debug(f"Highlight {idx}: similarity = {similarity:.3f}")
        
        if similarity >= threshold:
            start = match.a
            end = match.a + len(text)
            end = min(end, len(book_text))
            
            if not any(max(start, s) < min(end, e) for s, e in found_spans):
                found_spans.append((start, end))
                logger.debug(f"Match found with {similarity:.2%} similarity at position {start}-{end}")
    
    logger.debug(f"Difflib matching complete. Found {len(found_spans)} matches")
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
        logger.debug("punkt tokenizer found")
    except LookupError:
        logger.info("Downloading NLTK 'punkt' tokenizer...")
        nltk.download('punkt', quiet=True)
        logger.debug("punkt tokenizer downloaded")
    
    logger.info(f"Loading sentence transformer model: {NLP_MODEL}")
    model = SentenceTransformer(NLP_MODEL)
    logger.debug(f"Model loaded: {type(model)}")
    
    logger.info("Tokenizing book text into sentences...")
    book_sentences = nltk.sent_tokenize(book_text)
    logger.debug(f"Book split into {len(book_sentences)} sentences")
    
    if len(book_sentences) > 10000:
        logger.warning(f"Large number of sentences ({len(book_sentences)}). This may take a while...")
    
    logger.info("Encoding book sentences...")
    logger.debug(f"Encoding {len(book_sentences)} sentences with batch_size=32")
    book_embeddings = model.encode(
        book_sentences,
        convert_to_tensor=True,
        show_progress_bar=logger.level <= logging.INFO,
        normalize_embeddings=True,
        batch_size=32
    )
    logger.debug(f"Book embeddings shape: {book_embeddings.shape}")
    
    logger.info("Encoding highlights...")
    logger.debug(f"Encoding {len(highlights)} highlights")
    highlight_embeddings = model.encode(
        highlights,
        convert_to_tensor=True,
        show_progress_bar=logger.level <= logging.INFO,
        normalize_embeddings=True,
        batch_size=32
    )
    logger.debug(f"Highlight embeddings shape: {highlight_embeddings.shape}")

    logger.info("Computing similarity scores...")
    logger.debug("Running cosine similarity...")
    cosine_scores = util.cos_sim(highlight_embeddings, book_embeddings)
    logger.debug(f"Cosine scores shape: {cosine_scores.shape}")
    
    found_spans = []
    
    for i, text in enumerate(highlights):
        if i % 10 == 0 and logger.level <= logging.DEBUG:
            logger.debug(f"Processing highlight {i+1}/{len(highlights)}")
        
        best_match_idx = cosine_scores[i].argmax().item()
        confidence = cosine_scores[i][best_match_idx].item()
        
        if i < 5 or logger.level <= logging.DEBUG:
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
                    logger.debug(f"Match added at position {final_start}-{final_end}")
            else:
                logger.debug(f"Could not find sentence in book text")
        else:
            if i < 5:
                logger.debug(f"Highlight {i+1}: confidence too low ({confidence:.3f} < {threshold})")
    
    logger.debug(f"Vector matching complete. Found {len(found_spans)} matches")
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
        # Convert ebook - FIXED: was html_path, now html_file
        html_file, intermediate_file, extract_dir = convert_ebook_to_html(
            str(ebook_path), converter_path, output_format
        )
        
        # Extract title - FIXED: was html_path, now html_file
        logger.debug(f"Reading HTML file: {html_file}")
        with open(html_file, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f, 'html.parser')
            doc_title = soup.title.string.strip() if soup.title and soup.title.string else os.path.basename(ebook_path)
        
        logger.debug(f"Document title: {doc_title}")
        
        # Find highlights
        relevant_highlights = find_book_highlights(doc_title, all_clippings)
        
        if not relevant_highlights:
            result['error'] = "No highlights found"
            return result
        
        result['highlights_total'] = len(relevant_highlights)
        
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

        doc, found_count = create_highlighted_docx(
            html_file,
            relevant_highlights,
            doc_title,
            selected_matcher,
            threshold
        )
        
        result['highlights_found'] = found_count
        
        # Save document
        if not output_path:
            output_path = os.path.splitext(ebook_path)[0] + '_highlighted.docx'
        
        logger.debug(f"Output path: {output_path}")
        logger.info(f"💾 Saving document to: {output_path}")
        doc.save(output_path)
        logger.debug("Document saved successfully")
        
        result['success'] = True
        result['output_path'] = output_path
        
    except Exception as e:
        logger.error(f"Error processing book: {e}")
        logger.debug("Full exception:", exc_info=True)
        result['error'] = str(e)
    
    finally:
        # Cleanup
        if not keep_html:
            if html_file and os.path.exists(html_file):
                try:
                    os.remove(html_file)
                    logger.debug(f"Cleaned up: {os.path.basename(html_file)}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {html_file}: {e}")
            
            if extract_dir and os.path.exists(extract_dir):
                try:
                    shutil.rmtree(extract_dir)
                    logger.debug(f"Cleaned up extracted directory: {extract_dir}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {extract_dir}: {e}")
            
            if intermediate_file and os.path.exists(intermediate_file):
                try:
                    os.remove(intermediate_file)
                    logger.debug(f"Cleaned up: {os.path.basename(intermediate_file)}")
                except Exception as e:
                    logger.debug(f"Failed to clean up {intermediate_file}: {e}")
        else:
            logger.debug("Keeping intermediate files (--keep-html)")
    
    return result

# --- MAIN EXECUTION ---

def main():
    parser = argparse.ArgumentParser(
        description="Generate Word documents with Kindle highlights - WITH BATCH MODE",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Process ALL books with highlights (batch mode):
    python kindle_highlighter.py --library --clippings "My Clippings.txt" --batch -v
  
  Process only the best matching book:
    python kindle_highlighter.py --library --clippings "My Clippings.txt" -v
  
  Compare all methods on a single book:
    python kindle_highlighter.py --ebook book.mobi --clippings "My Clippings.txt" --compare -v
  
  List all books:
    python kindle_highlighter.py --list-books
        """
    )
    
    # Mode selection
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--library", action="store_true", 
                           help="Process book(s) from Calibre Library.")
    mode_group.add_argument("--list-books", action="store_true", 
                           help="List all books in Calibre Library and exit.")
    
    # Direct file mode
    direct_group = parser.add_argument_group('Direct File Mode')
    direct_group.add_argument("--ebook", dest="ebook_file", 
                             help="Path to a single ebook file.")
    direct_group.add_argument("--clippings", dest="clippings_file", 
                             help="Path to 'My Clippings.txt'.")
    
    # Library options
    library_group = parser.add_argument_group('Library Options')
    library_group.add_argument("--library-path", 
                              help=f"Path to Calibre Library (default: {DEFAULT_CALIBRE_LIBRARY})")
    library_group.add_argument("--batch", action="store_true",
                              help="Process ALL books with 95%+ match confidence")
    
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
    
    # Matching options
    matching_group = parser.add_argument_group('Matching Options')
    matching_group.add_argument(
        "-m", "--method",
        choices=['regex', 'diff', 'vector'],
        default='diff',
        help="Matching method: regex (fast), diff (balanced), vector (best accuracy)"
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
        
        if args.library and not clippings_path:
            logger.debug("Library mode: checking for default clippings file")
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
            pass  # Good!
        elif args.try_native:
            logger.warning("No standard converter found. Will try native conversion.")
        else:
            handle_error("No converter available. Install Calibre or use --try-native")
        
        # --- BATCH MODE (LIBRARY) ---
        if args.library:
            logger.debug("Library mode: auto-selecting book(s)")
            library_path = find_calibre_library(args.library_path)
            books = discover_books(library_path)
            clipping_titles = [c['title'] for c in all_clippings]
            
            # Match all books
            matches = match_books_to_clippings(books, clipping_titles, MIN_TITLE_MATCH_SCORE)
            
            if not matches:
                handle_error("No matching books found!")
            
            # Show matches
            logger.info("\n" + "=" * 80)
            logger.info("Book Matching Results:")
            logger.info("=" * 80)
            for idx, m in enumerate(matches[:20], 1):
                logger.info(f"{idx:2d}. Score {m['score']:3d}% | {m['highlight_count']:4d} highlights | '{m['book']['title']}'")
                logger.info(f"     Clipping title: '{m['clipping_title']}'")
            logger.info("=" * 80)
            
            # Determine which books to process
            if args.batch:
                # Process ALL books with 95%+ match
                books_to_process = [m for m in matches if m['score'] >= AUTO_SELECT_THRESHOLD]
                logger.info(f"\n📦 BATCH MODE: Processing {len(books_to_process)} books with {AUTO_SELECT_THRESHOLD}%+ match confidence")
            else:
                # Process only the best match
                if matches[0]['score'] >= AUTO_SELECT_THRESHOLD:
                    books_to_process = [matches[0]]
                    logger.info(f"\n✓ Auto-selected best match: '{matches[0]['book']['title']}' ({matches[0]['score']}%)")
                else:
                    logger.error(f"\n❌ Best match only {matches[0]['score']}% (need {AUTO_SELECT_THRESHOLD}%+)")
                    logger.error("   Use --batch to process all good matches, or specify with --ebook")
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
            if not ebook_path:
                handle_error("An ebook file must be provided with --ebook")
            
            logger.debug(f"Ebook path: {ebook_path}")
            
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
                pass
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