# MD 2 DOCX Converter

Convert all the MD files from a directory to DOCX file based on a template file.

## Usage

```console
python3 -B md2docx.py -c test/config.yaml
```

## Config File Format

The config file is in YAML format and has to contain the following Keys:
- **Folder**: where the MD files are located.
- **Filename**: the name of the result DOCX file.
- **Template**: the name of the template DOCX file.
- **Table Style**: the name of the table style it has to be the default table style in the document otherwise it will be not recognized.
- **Inline Code Font**: the font name of the inline code style.
- **Inline Code Color**: the color of the inline code style.
- **Syntax Highlight Style**: the style of the syntax highlighter: [link](https://pygments.org/demo/).

## Limitations

- Footnote is not supported
- Emojis is not supported
- Blockquotes is not supported
- Image title is not supported
- Text alignment in table is from word docx
- Subscript/Superscript is not supported
- Multiple horizontal rules under eachother is not supported
- Picture has its own paragraph
- List numbering is not reseting after the end of the list
- Table style has to be the default if you want to use it
- The following styles used by default from the DOCX template
    - The code block has to be in **code** style
    - The list styles has to be **ordered list<list level>** or **unordered list<list level>**

