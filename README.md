# NedanSukyan
Portable Invoice Price List Helper for IBM Palm Pilot/PalmOS 3.0</b>
================================================================
In my attempt to learn VB/VBA/HB, I've been developing this app for Palm OS 3.0 in "HB++". It is designed to function as portable application that can tell me exactly which steps I must take to check invoices for certain suppliers.</b>
I often find myself having to check which actions are appropriate for invoice checking, depending on the supplier. A lot of my notes on the differences between steps are held in my palm pilot, among other scattered notes; I thought it would be useful to combine VB learning with an application that would clean up my messy virtual sticky-notes.
# Main Supplier Screen
This screen gives you a non exhaustive list of suppliers of parts and materials to Euroglaze. The pages can be swiped through, and at the moment show 18 seperate supplier icons. Pressing on any of these will take you to the individual supplier information page for that company.</b>
![ScreenshotMainScreen](https://github.com/user-attachments/assets/c3540c34-679c-488f-a815-4aad1d9f21be)
# Individual Supplier Info Screen
Each supplier page will first include details about the suppliers location, and a button you can press to load the "Guide steps" for each respective company. The guide steps are the exact bulletpoints I use for remembering how to check invoices/price sheets for that supplier.</b>
![ScreenshotLiniar](https://github.com/user-attachments/assets/2d07c648-6e4d-41a0-b7fb-58aef04afaa8)
# About/Info Screen
Pressing the question mark button on the main page will take you to the "Help"/Info page. This area shows a little bit of background behind by the app exists, and a spinning 3D cube animation that I made by producing a function to quickly swap out frames of animation. Once the animation has played once, it can be played again by pressing the "Spin" button that appears. A large part of this section was understanding how dithering can be used to create realistic portrayals of 3D objects with monochrome graphics, and how quickly frames can be swapped in an animation before the CPU starts skipping. You have to remember we're running at about 16Mhz.</b>
![ScreenshotInforPage](https://github.com/user-attachments/assets/6ee21968-14a9-42bd-b651-0f17d8496286)
</b></b>
Once the program is fully finished I will move the compiled .prc file to a release section. There will also be two versions, one that is public facing that includes publically available information, and a Euroglaze private version. The private version is necessary as some of the information that I would like to be featured in the app (for my own personal use) will be sensitive data. </>
# Installation (Both current repo or releases)
The compiled .prc file provided will run natively on a PalmOS 3.0 or higher compatible device. I'm targetting 3.0 because it is the OS I'm using on my Palm Pilot III.</b>
Simply transfer the file through via hotsync, or Infrared Beam, and it should run natively. If you have a newer PalmOS handheld then you may be able to place this file onto an SD card and read from there.
