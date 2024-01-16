# Outlook Mails Word Cloud Generator

This Python script generates a word cloud from emails in a specified Outlook folder.

## Description

The `outlook-wordcloud-generator` script interfaces with Microsoft Outlook to extract emails from a specified folder and creates a word cloud from the text of these emails.

## Requirements

- Microsoft Outlook
- Python 3.x
- Python Libraries: `win32com.client`, `numpy`, `Pillow`, `wordcloud`, `matplotlib`, `nltk`

## Installation

Install the required Python libraries:

```bash
pip install pywin32 numpy Pillow wordcloud matplotlib nltk

```
## Usage

To use the Outlook Word Cloud Generator, follow these steps:

1. **Run the Script**: 
   Open a command line interface and navigate to the script's directory. Run the script using Python.
   ```bash
   # To generate directly from Windows outlook app folder
   python.exe wc_outlook_app.py
   ```
   ```bash
   # To generate from saved outlook messages in local folder (msg)
   python.exe wc_outlook.py
   ```
2. **Enter Outlook Folder Name**:
When prompted, enter the name of the Outlook folder from which you want to generate the word cloud.

3. **Word Cloud Generation**:
The script will generate a word cloud from the emails in the specified folder and save it as both a PNG and PDF file in the script's directory.

## Contributing

Contributions and collaborations to the Outlook Word Cloud Generator are welcome! Here's how you can contribute:

1. **Fork the Repository**: 
   Create your own fork of the project.

2. **Make Your Changes**: 
   Make modifications or add new features to your forked repository.

3. **Submit a Pull Request**: 
   Once you've made your changes, submit a pull request to the original repository for review.

For major changes or new features, please open an issue first to discuss what you would like to change. Please ensure to update tests as appropriate.
