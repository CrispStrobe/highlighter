# Kindle Ebook Highlighter

Automatically generate Word documents from your Kindle ebooks with all your highlights marked in yellow. 
Works with your Calibre library or individual ebook files.

## ✨ Features

- 📦 **Batch mode by default** - Processes ALL books from your Calibre library that match your highlights
- 📝 **Preserves formatting** - Bold, italic, underline, and paragraph structure maintained
- 🎯 **Three matching methods** - Choose between regex (fast), difflib (balanced), or vector (most accurate)
- 📚 **Calibre integration** - Automatically scans your library and matches books to highlights
- 🔍 **Fuzzy title matching** - Handles different editions and formatting variations
- 🌍 **Multi-language support** - Recognizes highlights in English, German, French, Spanish, and Italian
- 💾 **HTML preservation** - Optionally keep intermediate files for debugging
- 📊 **Detailed logging** - Verbose modes (-v, -vv) for troubleshooting

## 📋 Requirements

### Required Dependencies
```bash
pip install python-docx beautifulsoup4 thefuzz python-Levenshtein lxml
```

### Required Software
- **Calibre** (for ebook conversion) - Download from https://calibre-ebook.com/

### Optional Dependencies (for vector matching)
```bash
pip install sentence-transformers torch nltk
```

### Optional Dependencies (for native conversion)
```bash
pip install mobi ebooklib
```

## 🚀 Quick Start

### 1. Export Your Kindle Highlights
Connect your Kindle to your computer and copy the `My Clippings.txt` file from:
```
Kindle/documents/My Clippings.txt
```

### 2. Run the Script

**Default: Process ALL matched books from Calibre library**
```bash
python kindle_highlighter.py --clippings "My Clippings.txt" -v
```

**Process only the best matching book**
```bash
python kindle_highlighter.py --clippings "My Clippings.txt" --no-batch -v
```

**Process a specific ebook file**
```bash
python kindle_highlighter.py --ebook "book.mobi" --clippings "My Clippings.txt" -v
```

## 📖 Usage Examples

### List all books in your Calibre library
```bash
python kindle_highlighter.py --list-books
```

### Batch process with debugging output
```bash
python kindle_highlighter.py --clippings "My Clippings.txt" --preserve-html -vv
```

### Compare all three matching methods
```bash
python kindle_highlighter.py --ebook "book.mobi" --clippings "My Clippings.txt" --compare -v
```

### Use vector matching for best accuracy
```bash
python kindle_highlighter.py --clippings "My Clippings.txt" -m vector -v
```

### Specify custom Calibre library path
```bash
python kindle_highlighter.py --library-path "/path/to/library" --clippings "My Clippings.txt" -v
```

## 🎛️ Command Line Options

### Required
- `--clippings PATH` - Path to your "My Clippings.txt" file

### Mode Selection
- `--list-books` - List all books in Calibre library and exit
- `--ebook PATH` - Process a specific ebook file (overrides library mode)

### Library Options
- `--library-path PATH` - Custom Calibre library path (default: ~/Calibre Library/)
- `--no-batch` - Process only the best match instead of all books with 95%+ confidence

### Matching Options
- `-m, --method {regex,diff,vector}` - Matching method (default: diff)
  - `regex` - Fast, exact matches only
  - `diff` - Balanced, handles minor variations (default)
  - `vector` - Most accurate, uses AI semantic matching
- `--similarity-threshold FLOAT` - Threshold for difflib method (default: 0.9)
- `--vector-threshold FLOAT` - Threshold for vector method (default: 0.65)
- `--compare` - Run all three methods and compare results

### Output Options
- `-o, --output PATH` - Custom output path (single file mode only)
- `--preserve-html` - Keep intermediate HTML files for inspection

### Advanced Options
- `--calibre-path PATH` - Custom path to Calibre's ebook-convert binary
- `--try-native` - Attempt native Python conversion if Calibre fails
- `-v, --verbose` - Increase verbosity (-v for info, -vv for debug)

## 🔧 How It Works

1. **Scans** your Calibre library for all ebooks
2. **Parses** your My Clippings.txt file
3. **Matches** ebook titles to clipping titles using fuzzy matching
4. **Converts** ebooks to HTML format
5. **Extracts** text while preserving paragraph structure and formatting
6. **Finds** your highlights in the book text
7. **Generates** Word documents with highlights marked in yellow

## 📊 Matching Methods Explained

### Regex (Fast)
- Uses regular expressions for exact matching
- Handles minor spacing and hyphen variations
- Best for: Books with identical text to Kindle version

### Difflib (Balanced) - **DEFAULT**
- Uses Python's SequenceMatcher for fuzzy matching
- Adjustable similarity threshold
- Best for: Most books, handles minor formatting differences

### Vector (Most Accurate)
- Uses AI-powered semantic similarity
- Requires additional dependencies (sentence-transformers, torch)
- Best for: Books with significant reformatting or different editions

## 🎯 Matching Confidence

The script automatically matches books based on title similarity:
- **95-100%** - Excellent match, processed by default in batch mode
- **90-94%** - Good match, use `--no-batch` to process manually
- **85-89%** - Uncertain match, verify manually with `--ebook`
- **<85%** - Low confidence, not processed automatically

## 🐛 Troubleshooting

### No highlights found in output
1. Try different matching method: `-m vector`
2. Lower the similarity threshold: `--similarity-threshold 0.8`
3. Use `--preserve-html` to inspect the extracted text
4. Verify the ebook is not DRM-protected

### Book not matched to highlights
1. Check title format in Calibre vs. My Clippings.txt
2. Use `-vv` to see matching scores
3. Process manually with `--ebook` if automatic matching fails

### Calibre not found
1. Install Calibre: https://calibre-ebook.com/
2. Or specify path: `--calibre-path /path/to/ebook-convert`
3. Or use native conversion: `--try-native` (requires mobi/ebooklib)

### Paragraphs not preserved
1. Use `--preserve-html -vv` to inspect HTML structure
2. The script should automatically handle most HTML structures
3. If issues persist, file a bug report with sample HTML

## 📝 Output Format

Generated Word documents include:
- ✅ Book title as heading
- ✅ Proper paragraph breaks
- ✅ **Bold**, *italic*, and underlined text
- ✅ Highlights marked in yellow
- ✅ Clean, professional formatting

## 🔒 Privacy & DRM

This tool:
- ✅ Works with your own ebooks and highlights
- ✅ Processes everything locally on your computer
- ❌ Cannot process DRM-protected ebooks
- ❌ Does not remove DRM

## 💡 Tips

1. **Batch processing is default** - Run once to process all your highlighted books
2. **Use `-vv` for debugging** - See exactly what's happening at each step
3. **Keep HTML files** - Use `--preserve-html` to inspect extraction issues
4. **Try vector matching** - If default matching misses highlights, use `-m vector`
5. **Organize your library** - Good metadata in Calibre = better matching

## 🤝 Contributing

Issues and pull requests welcome! Please include:
- Sample (anonymized) My Clippings.txt entry
- Ebook format and source
- Command used and full output with `-vv`

## 📄 License

This script is provided as-is for personal use. License is MIT. Use responsibly and respect copyright laws.

## 🙏 Credits

Uses these excellent libraries:
- python-docx - Word document generation
- BeautifulSoup4 - HTML parsing
- thefuzz - Fuzzy string matching
- Calibre - Ebook conversion
- sentence-transformers - AI semantic matching (optional)
