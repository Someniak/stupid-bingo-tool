# 🎉 Bingo Card Generator - README 🎉

Hello and welcome to the most ridiculous readme file you've ever encountered! Here's everything you didn't want to know about the Bingo Card Generator:

## 📋 What is This Nonsense?

This Python script is pure magic. It reads a boring Excel file with participant names, bingo phrases, and questions; then it transmogrifies them into personalized Bingo cards. Almost like magic, but with way more bugs and frustration. 

## 📋 Files You (Unfortunately) Need

1. `bingo.xlsx` 
   - Must contain 'Naam', 'Bingo', and 'Vraag' columns because apparently, we love Dutch.

2. `Arial.ttf`
   - Find it somewhere in `/System/Library/Fonts/Supplemental/Arial.ttf`, assuming your system is as sane as your code.

3. This script (`export.py`)
   - You’re here, so you probably realized that already.

## 📋 How This Script Wastes Your Time

### 1. Nothing Works Before Loading Data
`load_data(file_path)` 
- Reads data pretending to be a secret agent. Validates columns and whines if it doesn't find them. 🍂

### 2. More Useless Grouping
`group_items_by_owner(bingo_items)` 
- Groups items like a neurotic squirrel hoarding acorns. 🥜

### 3. Generates Bingo Cards Like a Slot Machine
`generate_card_for_participant(participant, items_by_owner)`
- Randomly generates cards, hoping they'll fit into a neat 3x3 grid. 🎰

`generate_and_shuffle_bingo_cards(participants, items_by_owner)`
- Shuffles and hopes for the best. 🤞

### 4. Create Images Like Picasso (But Worse)
`wrap_text(text, font, max_width)`
- Attempts to wrap text, often failing miserably. 🖼️

`create_bingo_card_image(participant, card, font, title_font)`
- Makes a 3x3 grid pretending to be an artist. 🎨

`create_bingo_card_images_batch(...)`
- Creates batches because why not? 💼

### 5. Save Your Misery to Excel and Image Files
`saving_bingo_cards_to_excel(bingo_cards, output_file_path)`
- Saves bingo cards to Excel, feeding your spreadsheet addiction. 🧮

### 6. PDFs and ZIPs - Because We Love Complications
`images_to_pdf(image_folder, pdf_path)`
- Converts images to PDF, very tastefully over-complicating it. 📄

`compress_folder_to_zip(folder_path, zip_file_path)`
- Compresses the folder into a ZIP file because no one likes to have uncompressed fun. 📦

### 7. When It's All Over, Pray It Works
`main()`
- Where all the magic (or disaster) happens. Runs everything, creating a new Excel file, images, ZIP, and PDF. 🎉

## 📋 Mandatory Prerequisites

- **Python 3.6+** (Good luck with anything else)
- **Pandas**: pip install pandas
- **Pillow**: pip install pillow
- **ReportLab**: pip install reportlab
- **openpyxl**: pip install openpyxl

## 📋 How to Run This Madness

```sh
python export.py
```

## 📋 Enjoy the Chaos (or Not)

Seriously, run at your own risk. It might work perfectly, or it might cause you to pull out all of your hair. Either way, it's a ride. Buckle up! 🎢

Happy Bingo-ing! (Or not...)

---

*Disclaimer: The author takes no responsibility for any lost time, broken keyboards, or mild psychological trauma induced by using this script.*