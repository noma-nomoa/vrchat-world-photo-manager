# WorldShot Log

<p align="right">
  <a href="./README.md"><kbd>日本語</kbd></a>
  <a href="./README.en.md"><kbd>English</kbd></a>
  <a href="./README.ko.md"><kbd>한국어</kbd></a>
</p>

![WorldShot Log](./img/banner.png)

WorldShot Log is a Windows desktop app for organizing, revisiting, and preparing VRChat photos for sharing.

In addition to photo management by date and world, it includes image adjustments, cropping, AI subject selection, transparent-background export, privacy masking, text overlays, and image overlays. Original images are not overwritten; edited images are saved as separate files.

> WorldShot Log is not an official app from VRChat Inc.

## Main Features

### Photo Organization

- Drag and drop VRChat photos into the app
- Rescan registered folders
- Browse photos by year, month, and date
- Browse photos by world
- Filter by favorite status, orientation, labels, and World name
- Multi-select, Shift range selection, and drag selection
- Keyboard navigation in the gallery

### Photo Details

- View photos as cards
- Open a large preview in the photo details modal
- Move to previous and next photos
- Open the original image
- Open the containing folder
- Open the VRChat page
- Manage favorites, labels, and notes
- Manually edit World names and World URLs

### World Info Helpers

- Organize World IDs extracted from photos
- Automatically fetch World info after import
- Reuse fetched info for photos with the same World ID
- Display World names, descriptions, tags, and World URLs

### Built-in Image Editor

- Open the image editor from photo details
- Save edited images as separate files without overwriting originals
- Adjust brightness, exposure, contrast, highlights, shadows, whites, and blacks
- Adjust temperature, tint, saturation, and vibrance
- Adjust clarity, texture, fade, grain, and vignette
- Use smart auto correction, learned correction, and correction strength
- Edit RGB / HSV tone curves with histogram display
- Apply presets and save user presets
- Add transparent images as image overlays
- Manage, delete, and arrange multiple overlay layers
- Adjust opacity, blend mode, target area, layer order, and size
- Light snapping to rulers and grids

### Cropping and Composition

- Supports original, 1:1, 16:9, 9:16, 5:4, 4:5, 3:2, and 2:3 ratios
- Includes presets for VRC Gallery, emoji/stickers, and avatar thumbnails
- Keeps missing square-canvas areas transparent
- Resizes fixed-resolution presets to the intended output size when saving
- Adjust zoom, horizontal position, and vertical position
- Rotate 90 degrees, freely rotate, flip horizontally, and flip vertically
- Show rulers and a rule-of-thirds grid

### AI Subject Selection

- Generate local subject masks with a lightweight AI model
- Download standard AI withoutBG Snap and high-quality AI withoutBG Focus OSS only when needed
- Use generated subject masks as the target area for adjustments, blur, text, and image overlays
- Save transparent-background PNGs using the subject mask
- AI processing runs on your PC; images are not sent to external servers for AI processing
- Model and license details: [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md), [AI Model License Review](./docs/ai-model-license-review.md)

### Blur and Privacy Masking

- Full-image blur
- Radial blur
- Privacy masking with blur, mosaic, or fill
- Rectangle, circle, and freehand selections
- Move, resize, and rotate masks before applying
- Useful for hiding usernames, chat, UI, and personal information

### Text Overlays

- Add multiple text layers
- Change font size, font, weight, and color
- Add outlines, shadows, glow, and other styling
- Adjust letter spacing
- Move and rotate text
- Japanese-friendly fonts and decorative fonts
- Preview each font in the font selector
- Show recently used fonts

### Data Management

- Regenerate thumbnails
- Check for missing files, missing thumbnails, and photos without World info
- Create backups
- Restore from backups
- Export CSV / JSON
- In-app uninstall flow

## Use Cases

- Organizing VRChat photos
- Looking back on visited worlds
- Preparing images for X posts
- Preparing images for Booth listings
- Cropping thumbnails
- Hiding chat logs and usernames before posting
- Creating square images for VRC Gallery, emoji, and stickers

## Download

The latest version is available from GitHub Releases.

- [Releases](https://github.com/noma-nomoa/vrchat-world-photo-manager/releases)
- [v2.4.2 Release Notes](./release-notes/v2.4.2.md)
- [v2.4.1 Release Notes](./release-notes/v2.4.1.md)
- [v2.4.0 Release Notes](./release-notes/v2.4.0.md)
- [v2.3.0 Release Notes](./release-notes/v2.3.0.md)
- [v2.2.1 Release Notes](./release-notes/v2.2.1.md)
- [v2.2.0 Release Notes](./release-notes/v2.2.0.md)
- [v2.1.0 Release Notes](./release-notes/v2.1.0.md)
- [v2.0.0 Release Notes](./release-notes/v2.0.0.md)

The Windows installer is `WorldShotLogSetup.exe`.

## Installation

1. Download the latest `WorldShotLogSetup.exe` from GitHub Releases.
2. Run the downloaded `WorldShotLogSetup.exe`.
3. WorldShot Log starts after installation.

Windows may show a security warning. Please review it before running the installer.

## Updates

WorldShot Log supports update checks through GitHub Releases.

- The packaged app notifies you when a new version is available.
- After an update, the app shows the main changes on first launch.
- To update manually, download and run the latest `WorldShotLogSetup.exe` from GitHub Releases.
- In-app updates do not run in development mode (`npm start`).

## Privacy and Data Handling

- Imported photos, notes, labels, settings, and thumbnails are stored on your PC.
- Original images are not overwritten. Edited images are saved as separate files.
- The app may access external services such as VRChat and GitHub to fetch World info or check for updates.
- AI subject selection runs on your PC. External access is used only when downloading additional model files.
- Downloaded AI models and image overlay assets are stored in app-managed folders and can be deleted from inside the app.
- Backups and CSV / JSON exports are saved to locations selected by the user.
- If you choose to delete app data during uninstall, saved app data is removed.

## Data Locations

The app primarily stores data in the following locations.

- DB / settings: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\data\`
- AI models: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\models\`
- Image overlay assets: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\photo-editor-overlays\`
- Thumbnails: `C:\Users\<UserName>\WorldShot Log\thumbnails`
- Runtime cache: `C:\Users\<UserName>\AppData\Local\WorldShot Log\SessionData\`

## About This App

WorldShot Log is an independently created app. AI assistance is used for parts of design, implementation, and documentation organization. Final decisions on specifications, behavior, verification, and releases are made by the author.

## Related Documents

- [RELEASE.md](./RELEASE.md): Release workflow
- [AI_MAINTENANCE_GUIDE.md](./AI_MAINTENANCE_GUIDE.md): Maintenance and modification guide
- [release-notes/](./release-notes): Version history

## Limitations

- Automatic metadata fetching for private / non-public worlds is not supported.
- Choosing a custom install location is not supported.
- Distribution is currently intended for Windows only.

## License

WorldShot Log is released under the MIT License. See [LICENSE](./LICENSE) for details.
