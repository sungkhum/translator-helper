# translator-helper
This code utilizes ChatGPT to simplify and modernize a text to be used as an aid for translation. Input takes a Microsoft Word document and exports a Word document with the original text on the left and the modernized/simplified text on the right.

To run it, make sure you have Python 3 installed on your system. A tutorial on how to do so is here: https://kinsta.com/knowledgebase/install-python/#:~:text=Step%201%3A%20Download%20the%20Python,Python%20from%20the%20official%20website.

Download this code by clicking the "Code" button on the top right of this page and select "Download ZIP"

After unzipping the downloaded file, open up translation-helper.py in a text editor and add your ChatGPT API key to this line (line 8) preserving the single quotes around the key: openai.api_key = 'YOURAPIKEY'

To get an OpenAI ChatGPT API Key you can follow this tutorial: https://www.youtube.com/watch?v=aVog4J6nIAU

Then have your input.doc contain the text you want to modernize in perperation for translation in the same directory as translation-helper.py

And run the script in the terminal using: python3 translation-helper.py

A progress bar will appear if everything is working correctly, with an estimated wait time.

An average book will cost around $4-$10 because we are using ChatGPT-4 Turbo.
