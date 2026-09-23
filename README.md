# Guide to Converting Samsung Notes to Markdown files
This is a solution to converting Samsung Notes (both the legacy format and the new format) to the universal and superior markdown file type, intended for Obsidian (my new preferred note taking app). 
This converter also attempts to preserve timestamps to keep notes in order by the dates they were created and modified.
Additionally, if any note does not have a title, a title will be generated using either the first 3 words of the text or the first 15 characters. 

Fully transparency: I am not a programmer, this code was created using AI. 
A solution to my problem was not previously available, and now I want to share this tool with others who have 100's or 1000's of notes imprisoned on Samsung devices like I did.
If there are any issues with how I've set up Github or have any desire to reach out to me, you can email me at ScribeTheSol@gmail.com

There are two scripts which have different outputs, made for Windows OS. 

- Version 8 : Alongside creating titles for notes with and without titles, it also adds their creation date a the front of their title for sorting purposes.
Personally, I think this looks overwhelming and inefficient, however the option is available for those who want it for maybe another note-taking app.


- Version 9 : This removes the dates in the titles and adds sorting metadata into every note.

Every legacy file will have a metadata "tag:legacy-import" (which I believe is intended for sorting in Joplin).

Additionally, this requires creating a file/note in Obsidian to sort all your notes automatically. (I never got this feature to work properly, but it exists)

The file/note must be a markdown file type (.md) and named: sortspec

And it must contain the following text:

	---
	sorting-spec: |
  		target-folder: /*
  		< a-z by-metadata: created
	---

_______________________________________________________

Lastly, I have included a command line option to convert folders of notes as 'batches'. 
This helped me keep my notes more organized in Obsidian, because I couldn't get the sorting feature to work perfectly for the legacy notes.


  To use either of these scripts ;

- Download and Install Python if you haven't already: https://www.python.org/downloads/
(( I am using version 3.14.6 for this, but the latest should be fine ))

- Save a copy of the script you want to use onto your Desktop.

- Export all of your Samsung Notes files as their SDOCX file type into a folder on your Desktop.
(( you can do this by file sharing from your phone, or you can use the Samsung Notes app on the Microsoft Store ))

- Name the folder of all your SDOCX files you want to convert on your Desktop to: SDOCXfiles
- Create another folder in your Desktop named: MarkdownOutput 

-- If you want to convert your notes in batches, create folders (folder name should not matter) within your SDOCXfiles folder with up to 100 notes each,
and use the command that ends in --batch.

- The commands will prompt whether you want to change the creation dates of your legacy notes to a custom date or keep the originals ((helpful as an alternative sorting solution)).

___________

- Open Command Prompt as Administrator and run this command to change your directory first:
(( be sure to edit the command and add your computer's username ))

		cd C:\Users\*yourusername*\Desktop

- Next, choose which version of the script you want to use, and use the command with --batch at the end if you want to convert multiple folders of your SDOCX files at once. 

Version 8 :

	python sdocx_to_markdown8.py C:\Users\ryans\Desktop\SDOCXfiles C:\Users\ryans\Desktop\MarkdownOutput
or

	python sdocx_to_markdown8.py C:\Users\ryans\Desktop\SDOCXfiles C:\Users\ryans\Desktop\MarkdownOutput --batch


Version 9 :

	python sdocx_to_markdown9.py C:\Users\ryans\Desktop\SDOCXfiles C:\Users\ryans\Desktop\MarkdownOutput
or

	python sdocx_to_markdown9.py C:\Users\ryans\Desktop\SDOCXfiles C:\Users\ryans\Desktop\MarkdownOutput --batch

___________

Done! 
I hope this was as helpful for you as it was for me. 

Next I would look into how to set up an Obsidian Vault on your desktop and phone, if that's what you choose to do.

If you want to sync your Obsidian Notes between all your devices and never risk losing them, 
you can use the same guide I did to get that set up here: 
https://www.youtube.com/watch?v=ImrLbomFYA0





