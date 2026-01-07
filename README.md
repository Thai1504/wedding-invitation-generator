# Wedding Invitation Generator

A beautiful web-based wedding invitation generator for creating personalized wedding invitations with Vietnamese language support.

## Features

- 📝 **Excel/CSV Upload**: Upload guest lists from Excel or CSV files
- 🎨 **Elegant Design**: Two-column layout with couple photo and invitation details
- 🌏 **Vietnamese Support**: Full support for Vietnamese diacritics
- 📄 **PDF Export**: Download all invitations as PDFs organized by department
- 🖼️ **Image Export**: Save individual invitations as high-quality JPEG images
- 👥 **Batch Processing**: Generate invitations for hundreds of guests at once

## Demo

Visit the live demo: [Wedding Invitation Generator](https://[your-username].github.io/[repo-name]/)

## Usage

1. Open `index.html` in your web browser
2. Upload an Excel file with guest information (columns: "Mã nhân viên", "Kính gửi", "Phòng ban")
3. Preview the generated invitations
4. Download individual invitations or bulk download as a ZIP file

## Local Development

To run locally:

```bash
# Start a local server
./start-server.sh

# Or use Python
python3 -m http.server 8000

# Then open http://localhost:8000
```

## Customization

- **Couple Photo**: Replace `LJF00603.jpg` with your photo
- **Monogram**: Replace `IMG_1984.JPG` with your monogram
- **Fonts**: Add custom fonts to the `Fonts/` folder
- **Colors & Styling**: Edit CSS in `index.html`

## Tech Stack

- HTML5 / CSS3
- JavaScript (ES6+)
- Libraries:
  - html2canvas - For image generation
  - jsPDF - For PDF creation
  - JSZip - For batch downloads
  - SheetJS - For Excel file reading

## License

MIT License - Feel free to use for your own wedding!

## Credits

Created with ❤️ using [Claude Code](https://claude.com/claude-code)
