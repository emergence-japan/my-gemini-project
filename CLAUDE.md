# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains two static web projects:

1. **my-gemini-project**: AI bootcamp landing page promoting a community for learning generative AI
2. **test-project**: Single Mother Playful States Scale (SPSS) web application for psychological assessment

## Technology Stack

- **Frontend**: Pure HTML5, CSS3, and vanilla JavaScript
- **No build tools**: Projects are designed to run directly in browsers without compilation
- **No package management**: All dependencies are loaded via CDN links

## Development Commands

### Running the Projects
- Open `my-gemini-project/index.html` or `test-project/index.html` directly in a web browser
- Use a local development server for better experience:
  ```bash
  # If Python is available
  python -m http.server 8000
  
  # If Node.js is available
  npx serve .
  ```

### No Build Process
These are static HTML projects with no build, compile, or bundling steps required.

## Project Architecture

### my-gemini-project Structure
- `index.html`: Main landing page with comprehensive content about AI bootcamp
- `style.css`: Complete styling with responsive design, gradients, and animations
- `README.md`: Basic project title

**Key Features:**
- Responsive design with mobile-first approach
- Hero section with background images
- Grid layout for benefits and examples
- Interactive hover effects and animations
- Japanese language content focused on AI education

### test-project Structure
- `index.html`: Single-page psychological assessment application

**Key Features:**
- 25-question psychological survey with 5-point Likert scales
- Real-time progress tracking
- Automatic scoring across 5 factors:
  1. Daily joy discovery (日常の楽しさ発見)
  2. Freedom and liberation (自由感・解放感) 
  3. Creative and spontaneous activities (創造的・自発的活動)
  4. Playful interaction with children (子どもとのプレイフル交流)
  5. Social connection enjoyment (社会的つながりでの楽しさ)
- Data collection via Google Apps Script integration
- Results visualization with animated progress bars
- Print and copy functionality for results

## External Dependencies

### my-gemini-project
- Google Fonts: Noto Sans JP
- Unsplash images for hero background and profile photos

### test-project
- Embedded Google Apps Script URL for data collection: `https://script.google.com/macros/s/AKfycbxiNYS5UtOh9prjzIwEHgHDziC_F8tXvjMqCDpgyiapW0posq86cMYYsXXkCpH48WSZwA/exec`

## Testing

- **Manual testing**: Open files in different browsers and screen sizes
- **No automated tests**: Projects use simple DOM manipulation without test frameworks
- **Cross-browser compatibility**: Test in Chrome, Firefox, Safari, and Edge

## Development Notes

- Both projects use inline CSS and JavaScript for simplicity
- No framework dependencies or complex state management
- Files can be edited directly without any build process
- Images are loaded from external CDNs (Unsplash, placeholder services)