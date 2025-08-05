# PowerPoint Generator

A basic PowerPoint presentation generator that leverages AI and stock imagery to create presentations automatically.

## Features

- **AI-Powered Content Generation**: Uses Google's Gemini 2.5 Pro model to generate presentation content
- **Automatic Image Integration**: Fetches relevant images from Pexels to enhance slides
- **Automated Presentation Creation**: Generates complete PowerPoint files programmatically

## Prerequisites

- Python 3.7+
- Google AI Studio API key (for Gemini access)
- Pexels API key

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd powerpoint-generator
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Set up environment variables:
```bash
# Create a .env file in the project root
GEMINI_API_KEY=your_gemini_api_key_here
PEXELS_API_KEY=your_pexels_api_key_here
```

## Required Dependencies

```
google-generativeai
python-pptx
requests
python-dotenv
```

## Getting API Keys

### Gemini API Key
1. Visit [Google AI Studio](https://aistudio.google.com/)
2. Sign in with your Google account
3. Create a new API key
4. Copy the key to your `.env` file

### Pexels API Key
1. Visit [Pexels API](https://www.pexels.com/api/)
2. Sign up for a free account
3. Generate an API key
4. Copy the key to your `.env` file

## Usage

### Basic Usage

```python
from powerpoint_generator import PowerPointGenerator

# Initialize the generator
generator = PowerPointGenerator()

# Generate a presentation
presentation = generator.create_presentation(
    topic="Climate Change",
    num_slides=5
)

# Save the presentation
presentation.save("climate_change_presentation.pptx")
```

### Command Line Usage

```bash
python generate_presentation.py --topic "Your Topic" --slides 5 --output "presentation.pptx"
```

## Project Structure

```
powerpoint-generator/
├── README.md
├── requirements.txt
├── .env.example
├── generate_presentation.py
├── powerpoint_generator/
│   ├── __init__.py
│   ├── generator.py
│   ├── gemini_client.py
│   └── pexels_client.py
└── examples/
    └── sample_presentation.pptx
```

## Configuration

The application uses the following default settings:

- **Gemini Model**: gemini-2.0-flash-exp (or gemini-1.5-pro as fallback)
- **Image Resolution**: 1920x1080 (landscape)
- **Max Images per Slide**: 1
- **Content Language**: English

## API Limits

### Gemini API
- Free tier: 15 requests per minute
- Rate limiting is handled automatically

### Pexels API
- Free tier: 200 requests per hour
- 20,000 requests per month

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Troubleshooting

### Common Issues

**API Key Errors**
- Ensure your API keys are correctly set in the `.env` file
- Verify that your API keys are active and have sufficient quota

**Image Download Failures**
- Check your internet connection
- Verify Pexels API key is valid
- Some images may not be available in the requested resolution

**PowerPoint Generation Errors**
- Ensure you have write permissions in the output directory
- Check that all required dependencies are installed

## Acknowledgments

- [Google Gemini](https://deepmind.google/technologies/gemini/) for AI content generation
- [Pexels](https://www.pexels.com/) for providing free stock photography
- [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint file manipulation

## Support

For issues and questions, please open an issue on the GitHub repository or contact the maintainers.
