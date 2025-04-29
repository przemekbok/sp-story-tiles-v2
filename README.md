# Content Gallery Web Part for SharePoint

A modern SharePoint Framework (SPFx) web part that displays content in a tile-based layout with modal functionality for detailed information.

## Features

- Displays 1-4 content tiles per view with configurable layout
- Carousel navigation for browsing through additional content items
- Pulls content dynamically from a SharePoint list with images as attachments
- Modal window shows detailed content information when a tile is clicked
- Direct link support for tiles with URLs
- Responsive design that works across all device sizes
- Modern UI with shadows, rounded corners, and hover effects
- Configurable fields for customization

## Getting Started

### Prerequisites

- Node.js (version 18.17.1 or higher)
- SharePoint Developer environment
- SharePoint list with the correct content type (see below)

### Installation

1. Clone this repository
2. Run `npm install`
3. Run `gulp serve` to test locally
4. Run `gulp bundle --ship` and `gulp package-solution --ship` to package for deployment
5. Upload the `.sppkg` file from the `sharepoint/solution` folder to your SharePoint App Catalog
6. Add the web part to your page

## SharePoint List Setup

### Content Type Columns

Create a SharePoint list with the following columns:

1. **Title** (Default column)
   - Used for the tile title

2. **Description**
   - Type: Multiple lines of text
   - Used for the content's description in the tile and modal

3. **Image**
   - Type: Thumbnail
   - Used to display an image for the content
   - Stores images as attachments

4. **LinkUrl**
   - Type: Single line of text
   - Optional URL that will open when clicking the tile
   - If empty, clicking will open the modal instead

### Creating the List

1. Create a new list in SharePoint (suggested name: "ContentLibrary")
2. Add the columns specified above
3. Add your content items with images
4. Configure the web part to use this list

## Web Part Configuration

In the web part properties pane, you can configure:

- **Web Part Title**: The heading displayed above the tiles
- **SharePoint List Name**: Name of the list containing your content
- **Tiles Per View**: Number of tiles to display at once (1-4)
- **Field Name Settings**: Configure custom field names if they differ from defaults

## Technical Details

### PnP JS Integration

This web part uses PnP JS (SharePoint Patterns and Practices JavaScript library) for SharePoint data operations, which offers several advantages:

- Cleaner, more maintainable code for SharePoint operations
- Improved error handling and fallback mechanisms
- Better performance through optimized queries

### Image Handling

The web part implements a multi-layered approach to retrieve images:

1. **Basic Information**: Gets content title and description from the SharePoint list
2. **Image Retrieval**: Retrieves image attachments from the SharePoint list
3. **Fallback Image**: Uses a default image if no image is found

### Modal Implementation

The modal dialog is implemented using Fluent UI components:

1. **Responsive Design**: Works well on all screen sizes
2. **Accessible**: Implements proper keyboard navigation and accessibility features
3. **Themed**: Supports SharePoint themes including dark mode

## Development Notes

- Built using SharePoint Framework (SPFx) 1.20.0
- Uses React and Fluent UI components
- Implements responsive grid layout with CSS Grid
- Leverages PnP JS library for enhanced SharePoint operations