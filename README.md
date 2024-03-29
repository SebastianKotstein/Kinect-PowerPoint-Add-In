# Kinect-PowerPoint-Add-In
PowerPoint VSTO Add-in enabling users to control a slideshow with Kinect gestures and (optionally) with speech commands. 

## Hardware and Software Requirements:
To use this VSTO Add-in, you need the following hardware/software:
* A machine with a USB 3.0 port/chipset that is compatible with the Kinect Adapter for Windows
* A Kinect 2.0 sensor and a Kinect Adapter for Windows
* Windows 10, PowerPoint, and the Kinect Driver installed (download and install the official [Kinect 2.0 SDK](https://www.microsoft.com/en-us/download/details.aspx?id=44561))

## Supported Kinect Gestures:
After enabling gesture recognition over the Ribbon menu (labeled as "Kinect Gesture Control"), the Add-in recognizes the following gestures:
* Close your left hand to start the slide show
* Close your right hand to end the slide show 
* Swipe your hand from the right side of your body to the left to proceed to the next slide
* Rotate your arm counterclockwise to go back to the previous slide

## Supported Speech Commands:
The current version of this Add-in supports speech commands in English and German.
The set of usable speech recognizers and voice synthesizers, however, depends on your machine's culture and installed Windows 10 language packs.
To enable speech recognition, select an input language/culture in the combo box labeled "Input Language".
For English, the following speech commands are available:
* Start Recognition
* End Recognition
* Start Slide Show
* End Slide Show
* Am I tracked?
* What can I say?

For German, the following commands are recognized:
* Starte Gestensteuerung
* Beende Gestensteuerung
* Starte Präsentation
* Beende Präsentation
* Werde ich erkannt?
* Was kann ich sagen?

Commands for additional languages have not been added yet, but we have prepared the code implementation for Spanish, French, and Italian.




