import pandas as pd
import random
import os
import shutil
from collections import defaultdict
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def load_data(file_path):
    df = pd.read_excel(file_path)
    if 'Naam' not in df.columns or 'Bingo' not in df.columns or 'Vraag' not in df.columns:
        raise ValueError("Excel file must contain 'Naam', 'Bingo' and 'Vraag' columns.")
    df = df[['Naam', 'Bingo', 'Vraag']]
    participants = df['Naam'].unique().tolist()
    bingo_items = df.to_dict('records')
    
    for item in bingo_items:
        if 'Vraag' not in item:
            raise ValueError(f"Item missing 'Vraag' key: {item}")
    
    return participants, bingo_items

def group_items_by_owner(bingo_items):
    items_by_owner = defaultdict(list)
    for item in bingo_items:
        items_by_owner[item['Naam']].append(item)
    return items_by_owner

def generate_card_for_participant(participant, items_by_owner):
    card = []
    owners_used = set()
    used_questions = set()

    while len(card) < 9 and len(owners_used) < 3:
        available_owners = [owner for owner in items_by_owner.keys() if owner != participant and owner not in owners_used]
        if not available_owners:
            break
        
        random_owner = random.choice(available_owners)
        items_from_owner = items_by_owner[random_owner]
        random.shuffle(items_from_owner)
        
        count = 0
        for item in items_from_owner:
            if len(card) < 9 and count < 3 and item['Vraag'] not in used_questions:
                card.append(item)
                used_questions.add(item['Vraag'])
                count += 1
            if count == 3:
                break
        
        owners_used.add(random_owner)
    
    if len(owners_used) != 3 or len(card) != 9:
        return None
    
    random.shuffle(card)
    return card

def generate_and_shuffle_bingo_cards(participants, items_by_owner):
    cards = {}
    for participant in participants:
        card = None
        while card is None:
            card = generate_card_for_participant(participant, items_by_owner)
        cards[participant] = card
    return cards

def wrap_text(text, font, max_width):
    lines = []
    words = text.split()
    while words:
        line = ''
        while words and font.getlength(line + words[0]) <= max_width:
            line = line + (words.pop(0) + ' ')
        lines.append(line.strip())
    return lines

def create_bingo_card_image(participant, card, font, title_font):
    img = Image.new('RGB', (800, 800), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    # Add title
    title_width, title_height = d.textbbox((0, 0), participant, font=title_font)[2:4]
    d.text((400 - title_width / 2, 50 - title_height / 2), f"{participant}", fill=(0, 0, 0), font=title_font)
    
    # Draw 3x3 grid
    grid_size = 3
    cell_size = 200
    start_x, start_y = 100, 100

    for i in range(grid_size + 1):
        d.line((start_x, start_y + i * cell_size, start_x + grid_size * cell_size, start_y + i * cell_size), fill=(0, 0, 0), width=2)
        d.line((start_x + i * cell_size, start_y, start_x + i * cell_size, start_y + grid_size * cell_size), fill=(0, 0, 0), width=2)

    # Add text to grid
    for i, item in enumerate(card):
        row = i // grid_size
        col = i % grid_size
        text_x = start_x + col * cell_size
        text_y = start_y + row * cell_size
        wrapped_text = wrap_text(item['Bingo'], font, cell_size - 20)
        
        # Calculate the total height of the text block
        total_text_height = len(wrapped_text) * 24  # Assuming line height of 24
        start_text_y = text_y + (cell_size - total_text_height) / 2  # Center vertically
        
        for j, line in enumerate(wrapped_text):
            line_width = font.getlength(line)
            d.text((text_x + (cell_size - line_width) / 2, start_text_y + j * 24), line, fill=(0, 0, 0), font=font)  # Center horizontally
        
    return img



def create_bingo_card_images_batch(bingo_cards, output_folder, start_index, end_index, font, title_font):
    participants = list(bingo_cards.keys())
    batch = participants[start_index:end_index]
    for participant in batch:
        card = bingo_cards[participant]
        img = create_bingo_card_image(participant, card, font, title_font)
        img.save(f"{output_folder}/{participant}__bingo_card.png")

def save_bingo_cards_to_excel(bingo_cards, output_file_path):
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for participant, card in bingo_cards.items():
            card_df = pd.DataFrame(card)
            card_df['Deelnemer'] = participant
            card_df.to_excel(writer, sheet_name=participant, index=False)

def images_to_pdf(image_folder, pdf_path):
    images = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith('.png')]
    images.sort()

    c = canvas.Canvas(pdf_path, pagesize=A4)
    a4_width, a4_height = A4

    for img_path in images:
        img = Image.open(img_path)
        img_width, img_height = img.size

        # Calculate the resize factor for the image to fit on A4 while maintaining aspect ratio
        factor = min(a4_width / img_width, a4_height / img_height)
        resized_width = img_width * factor
        resized_height = img_height * factor

        # Calculate the position to center the image
        x = (a4_width - resized_width) / 2
        y = (a4_height - resized_height) / 2

        c.drawImage(img_path, x, y, width=resized_width, height=resized_height)
        c.showPage()

    c.save()
def create_and_save_bingo_card_images(bingo_cards, output_folder, font_path, batch_size=5):
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        font = ImageFont.truetype(font_path, 18)
        title_font = ImageFont.truetype(font_path, 36)
    except IOError:
        font = ImageFont.load_default()
        title_font = ImageFont.load_default()
    
    for i in range(0, len(bingo_cards), batch_size):
        create_bingo_card_images_batch(bingo_cards, output_folder, i, i + batch_size, font, title_font)

def compress_folder_to_zip(folder_path, zip_file_path):
    shutil.make_archive(zip_file_path, 'zip', folder_path)

def main():
    file_path = 'bingo.xlsx'  # Update this path to your bingo.xlsx file
    participants, bingo_items = load_data(file_path)
    items_by_owner = group_items_by_owner(bingo_items)

    # Generate the bingo cards
    bingo_cards = generate_and_shuffle_bingo_cards(participants, items_by_owner)
    
    # Save the bingo cards to a new Excel file
    output_file_path_linked = 'bingo_cards_output_linked.xlsx'
    save_bingo_cards_to_excel(bingo_cards, output_file_path_linked)
    
    # Save the bingo card images to a folder
    output_folder_linked = 'bingo_cards_images_linked'
    font_path = "/System/Library/Fonts/Supplemental/Arial.ttf"  # Update this path to your font file if required
    create_and_save_bingo_card_images(bingo_cards, output_folder_linked, font_path)
    
    # Compress the output folder into a zip file
    output_zip_path_linked = 'bingo_cards_images_linked'
    compress_folder_to_zip(output_folder_linked, output_zip_path_linked)
    
    # Convert images to PDF
    pdf_file_path = 'bingo_cards.pdf'
    images_to_pdf(output_folder_linked, pdf_file_path)

    print(f"Bingo cards Excel file saved to: {output_file_path_linked}")
    print(f"Bingo cards images zip file saved to: {output_zip_path_linked}.zip")

if __name__ == "__main__":
    main()
