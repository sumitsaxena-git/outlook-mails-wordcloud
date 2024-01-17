import os
import win32com.client
import numpy as np
from PIL import Image
from wordcloud import WordCloud
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import nltk
nltk.download('stopwords')
from nltk.corpus import stopwords

def get_outlook_messages(folder_name):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the index of Inbox
    target_folder = inbox.Folders[folder_name] if folder_name in [f.Name for f in inbox.Folders] else inbox

    messages = target_folder.Items
    message_text = " "
    for message in messages:
        try:
            message_text += message.Body
        except Exception as e:
            print("Error reading message: ", e)
    return message_text

script_directory = os.getcwd()

# Replace 'YourFolderName' with the name of the folder you want to read from
# folder_name = 'Secure'  # Replace with your folder name
folder_name = input('Enter the outlook folder name : ') # Replace with your folder name
msg_message = get_outlook_messages(folder_name)

wc = WordCloud(
    background_color='white',
    #colormap='binary',
    colormap='jet',
    stopwords=set(stopwords.words('english')),
    width=800,
    height=600
).generate(msg_message)

plt.axis("off")
plt.imshow(wc)

# Save the WordCloud as an image (e.g., PNG)
output_path = os.path.join(script_directory, 'wordcloud_outlook.png')
plt.savefig(output_path, bbox_inches='tight', pad_inches=0, dpi=600)

# Save the WordCloud as a PDF
output_path = os.path.join(script_directory, 'wordcloud_outlook.pdf')
plt.savefig(output_path, format='pdf', bbox_inches='tight', pad_inches=0)

print(f"WordCloud saved to : {script_directory} as - wordcloud_outlook file in pdf and png format")