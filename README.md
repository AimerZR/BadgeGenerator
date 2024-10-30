# BadgeGenerator

This repository provides a Python-based solution for generating badges in PDF format, based on a customizable configuration. The tool leverages settings defined in `config.json` for dynamic badge creation, allowing users to specify design presets, text placements, font sizes, colors, and other layout details for both the front and back of badges.

## Features

- **Preset-based Layouts**: Define various badge presets, including front and back designs, with specific text, font, color, and positioning.
- **Customizable Badge Elements**: Specify element positions for names, IDs, departments, and photos using JSON configuration.
- **PDF Output**: Generates badges directly into a multi-page PDF document.
- **Error Logging**: Records any errors encountered during badge generation to a log file.

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/BadgeGenerator.git
    ```
2. Ensure Python 3.x is installed.
3. Install required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place font and background image files as specified in `config.json`.
2. Customize `config.json` for badge parameters and layouts. Each preset should specify:
   - Badge layout for front and back
   - Positioning for text fields, photo, font type, and color.
   - Size and color configurations for each text element.

3. Run the script:
    ```bash
    python BadgeGenerator.py
    ```

4. Generated badges will be saved in the folder specified in `config.json`, defaulting to `Badge_output`.

## Configuration

### config.json
The `config.json` file defines layout options, paths, and other badge generation settings. Key parameters:

- **badge_folder**: Output directory for badge files.
- **badge_front_prefix / badge_back_prefix**: Filename prefix for front and back badge images.
- **output_pdf**: Destination for the final PDF file.
- **error_log**: File for recording errors.
- **presets**: Badge design presets, with options for front and back designs.

Each preset includes specifications for:
  - **Positioning**: Coordinates for text and photo elements.
  - **Font Settings**: Fonts for various badge elements (name, department, etc.).
  - **Colors and Sizes**: Defines color codes and text sizes.

#### Example Configuration Snippet
```json
{
  "badge_folder": "Badge_output",
  "badge_front_prefix": "badge_front_",
  "badge_back_prefix": "badge_back_",
  "output_pdf": "print/badge_output.pdf",
  "presets": {
    "ExamplePreset": {
      "front": {
        "name_char_limit": 12,
        "background_img": "bg/FBG.jpg",
        "font_name": "font/OpenSans-Bold.ttf",
        "photo_position": [80, 490],
        "name_position": [85, 90],
        "name_size": 64,
        "name_color": "#FFFFFF"
      },
      "back": {
        "background_img": "bg/BBG.jpg",
        "font_name": "font/OpenSans-Bold.ttf",
        "name_position": [80, 100],
        "name_size": 64,
        "name_color": "#000000"
      }
    }
  }
}
