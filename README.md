# Randomized Greeting and Quote VBS Script for Outlook Desktop
This is a super simple Visual Basic Script (VBS) that replaces placeholder text in your signature with a randomized **greeting** and **quote** by selecting a random line from a .txt document stored on your computer. 

## How it works
The script injects your HTML signature at the bottom of an email when you send it.
It replaces the placeholdertext `[[RANDOM_GREETING]]` and `[[RANDOM_QUOTE]]` already present in your signature with a random line from two separate .txt files you specify.

## Prerequisites
- You need a HTML signature saved as an .HTML file. Export your already existing signature into .HTML or download and customize the `signature_template.html` file provided above.
- This script is only applicable to the desktop version of Outlook.
- You need to have access to the VBA Editor in Outlook, some workplaces or schools limit the access to these due to security risks.

## Script Setup
1. Download the `outlook_randomizer.vbs` script from this repository.
2. Create two new text files somewhere on your computer and name them `greetings.txt` and the other `quotes.txt`.
3. Add lines of your desired text to both text files, each separated per line.   
![alt text](https://i.imgur.com/wMfMj8G.png)
5. Open the `outlookrandomizer.vbs` script using a text editor. (Right-click > Open With > Notepad)
6. Replace the placeholder `C:\path\to\greetings.txt` with the actual path to your greetings.txt file.    
Example: *C:\Users\BACONGUDEN\Documents\OutlookRandomizer\greetings.txt*   
8. Replace the placeholder `C:\path\to\quotes.txt` with the actual path to your quotes.txt file.    
Example: *C:\Users\BACONGUDEN\Documents\OutlookRandomizer\quotes.txt*   
9. Replace the placeholder `C:\path\to\signature.html` with the actual path to your signature.html file.    
Example: *C:\Users\BACONGUDEN\Documents\OutlookRandomizer\signature.html*   
10. Open Outlook, and press Alt + F11 to open upp the Outlook Script Editor.
11. Go to `Tools > References` and enable `Microsoft Word Object Library` and click OK.
12. Open `Project1` on the left-hand side, as well as `Microsoft Outlook Objects`.
13. Double-click `ThisOutlookSession` and copy-paste the code from `outlook_randomizer.vbs` into that window.
14. Press **CTRL + S** on your keyboard to save, alternatively press `File > Save VbaProject.OTM`.

## Finishing Setup
- Make sure that you have the placeholder texts `[[RANDOM_GREETING]]` and `[[RANDOM_QUOTE]]` present where you want them in the `signature.html` file you specified on step 9. There is a template HTML signature provided in the project. See `signature_template.html`.
- Remove any signature you have already in Outlook. (Outlook > File > Options > Mail > Signatures... > Set both dropdowns under "Choose default signature" to (none) and click "OK".
- Test the script by sending an email to yourself.

## Troubleshooting
**Q:** "User-defined type not defined" error, help.    
**A:** See Step 11 of Script Setup.

**Q:** I'm getting a security warning preventing the script from running.  
**A:** In order to use this script, either digitally sign in yourself or ensure that you have turned on "Enable all macros" in Outlook. Outlook's standard security settings restrict certain code execution to prevent security risks. If you wish to bypass the setting, refer to the "Disclaimer and Warnings" section.

*Any other troubleshooting, open an issue.*

## License
This project is licensed under the [MIT License](https://pastebin.com/yxBL3p16).

## Contributions
Contributions are welcome. I have very basic knowledge of VBS so if you have any suggestions, improvements, or bug fixes, feel free to open an issue or submit a pull request.

## Disclaimer and Warnings
> Please use this script responsibly and respect the privacy and legal restrictions of others when using it.    
> In order to use this script, either digitally sign in yourself or ensure that you have turned on "Enable all macros" in Outlook.    
> **Warning:** Enabling the "Enable all macros" option allows all scripts to run without warning, which may result in the execution of potentially malicious code. Exercise caution when enabling this setting.    
> To do so, navigate to Outlook > File > Options > Trust Center > Macro Settings and select "Enable all macros".    
> The author takes no responsibility for any misuse or damage caused by this script.   
