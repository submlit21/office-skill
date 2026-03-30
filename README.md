# Office Skill

Office document processing tool using `cli-anything-libreoffice` backend.

## Dependencies

- **Python**: >= 3.8
- **Core**:
  - `cli-anything-libreoffice>=0.1.0` - Backend Office document processing
  - `pyyaml>=6.0` - YAML configuration parsing
- **Templating**:
  - `jinja2>=3.0.0` - Advanced template variable substitution
  - `markupsafe>=2.0.0` - Jinja2 security dependency

## Installation

### From source (development)
```bash
# Clone repository
git clone https://github.com/example/office-skill.git
cd office-skill

# Install in development mode
pip install -e .
```

### Optional dependencies for enhanced functionality
- `pandoc` - Better DOCX to Markdown conversion (optional, system package)
- `libreoffice` - Alternative document conversion (optional, system package)
