#pip install extract_msg numpy Pillow wordcloud matplotlib nltk
# python.exe wc_outlook.py

import os
import extract_msg
import numpy as np
from PIL import Image
from wordcloud import WordCloud
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from nltk.corpus import stopwords
from os import path
from glob import glob

def find_ext(dr, ext):
    return glob(path.join(dr, "*.{}".format(ext)))

# Get the current working directory
script_directory = os.getcwd()

# Combine the script directory with the relative path to your Outlook message files
outlook_messages_path = os.path.join(script_directory, 'msg')
# print(outlook_messages_path)

# Validate the path
if not os.path.exists(outlook_messages_path):
    print("Invalid path. Please provide a valid path.")
    exit()

file_list = find_ext(outlook_messages_path, 'msg')

msg_message = ' '
for f in file_list:
    msg = extract_msg.Message(f)
    msg_message = msg_message + msg.body
    msg.close()

wc = WordCloud(
    background_color='white',
    #colormap='binary',
    colormap='jet',
    stopwords=['google'],
    width=800,
    height=600
).generate(msg_message)

plt.axis("off")
plt.imshow(wc)
plt.show()

# Save the WordCloud as an image (e.g., PNG)
output_path = os.path.join(script_directory, 'wordcloud.png')
plt.savefig(output_path, bbox_inches='tight', pad_inches=0, dpi=600)


# Save the WordCloud as a PDF
output_path = os.path.join(script_directory, 'wordcloud.pdf')
plt.savefig(output_path, format='pdf', bbox_inches='tight', pad_inches=0)

print(f"WordCloud saved as: {output_path}")