# Colab2Word

A Python package for generating professional Word documents from Google Colab notebooks.

## Installation

```bash
pip install git+https://github.com/nativocloud/Colab2Word.git
```

In Google Colab:
```python
!pip install git+https://github.com/nativocloud/Colab2Word.git
```

## Quick Start

```python
from colab2word import DocumentGenerator, DocumentSection

# Initialize document generator
doc_gen = DocumentGenerator(
    output_dir='output',
    default_filename='my_report.docx'
)

# Create a section
section = DocumentSection(
    content="Hello World!",
    content_type="text",
    header="Introduction"
)

# Generate document
doc_gen.generate_document(
    sections=[section],
    title="My First Document",
    author="Your Name"
)
```

## Features

- Generate professional Word documents from Colab notebooks
- Support for text, tables, and plots
- Customizable themes and styling
- Automatic content type detection
- Professional formatting and layout

## License

MIT License# Colab2Word

A Python package for generating professional Word documents from Google Colab notebooks.

## Installation

```bash
pip install git+https://github.com/nativocloud/Colab2Word.git
```

In Google Colab:
```python
!pip install git+https://github.com/nativocloud/Colab2Word.git
```

## Quick Start

```python
from colab2word import DocumentGenerator, DocumentSection

# Initialize document generator
doc_gen = DocumentGenerator(
    output_dir='output',
    default_filename='my_report.docx'
)

# Create a section
section = DocumentSection(
    content="Hello World!",
    content_type="text",
    header="Introduction"
)

# Generate document
doc_gen.generate_document(
    sections=[section],
    title="My First Document",
    author="Your Name"
)
```

## Features

- Generate professional Word documents from Colab notebooks
- Support for text, tables, and plots
- Customizable themes and styling
- Automatic content type detection
- Professional formatting and layout

## License

MIT License