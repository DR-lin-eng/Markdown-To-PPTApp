# Markdown to PowerPoint Converter

This Python application converts Markdown text into a PowerPoint presentation, allowing users to incorporate custom background images and formatted text into slides.

## Features

- **Select Background Image**: Users can choose a custom background image for all slides in the presentation.
- **Import Markdown File**: Supports importing Markdown files for conversion.
- **Convert to PowerPoint**: Converts the imported or pasted Markdown content into a PowerPoint file with formatting and styling.

## Requirements

- Python 3.6+
- wxPython
- python-pptx

## Installation

1. **Install Python**: Ensure Python 3.6 or newer is installed on your system.
2. **Install Dependencies**:
   ```bash
   pip install wxPython python-pptx
   ```

## Usage

1. **Run the Application**:
   ```bash
   python markdown_to_ppt.py
   ```
2. **Select a Background Image**: Click on '选择底图' and choose an image file.
3. **Import Markdown File**: Click on '导入Markdown文件' to load your Markdown content.
4. **Convert to PowerPoint**: After loading the content, click '转换为PPT' to generate the PowerPoint presentation.

## Application Structure

- `MarkdownToPPTApp`: Main application window handling UI and user interactions.
- `on_select_bg_image`: Function to select and set the background image for slides.
- `on_load_markdown`: Function to load Markdown content from a file.
- `on_convert`: Function that orchestrates the conversion process from Markdown to PowerPoint.
- `convert_markdown_to_ppt`: Core logic for creating PowerPoint slides based on Markdown content.
- `set_background_image`, `handle_slide_data`, `add_title`, `add_content`: Helper functions for slide creation and formatting.

## Limitations

- The application currently only supports basic text and title formatting based on specific Markdown cues.
- More complex Markdown features like tables, code syntax highlighting, and nested lists may not be rendered accurately.

## Contributing

Contributions to enhance the functionality, improve user interface, or extend compatibility are welcome. Please fork the repository, make your changes, and submit a pull request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
```
