# briefing_generator

Python script to generate morning weather briefings for Olympex.

Requires:
  --python-pptx (https://python-pptx.readthedocs.org/en/latest/)
  Follow instructions on that site for how to install this library to your python distribution.  We've tested both the pip and easy_install (for Mac) methods and they both seem to work.

To run the script:

--Place the python script in a working directory where you want your presentations to be generated

--Either load the script in your favorite python editor (like Spyder or IDLE) or open a console/terminal/powershell and go to the directory where the script is located.

--Run the script from the python editor or, if in a console, type "python presentation_maker.py" and hit enter

--The script should loop through and look for images to populate the presentation.  It will make a new subdirectory in your working directory with the date of the presentation.  All the images in the presentation and the .pptx file itself will be placed in that directory.

--Without any editing, the script is set to generate a "morning" briefing for the current day and will just use the current date at 12Z as the base time.
  
