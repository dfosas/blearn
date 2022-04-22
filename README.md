# blearn

[![Code style: black](
https://img.shields.io/badge/code%20style-black-000000.svg)](
https://github.com/psf/black)

Utilities to work with BlackBoard Learn in offline mode (this is, *not* with its API).

## TODO regarding code
* Remove print statements and `verbose` arguments and move to proper logging system.
  * Provide coding snippet to show how to use logging system or create it by default.

# Functionality
* Parse offline grading sheet.
* Parse bundled submission in zip file or extracted zip file:
  * Metadata based on file names.
  * Metadata from the content in the log `.txt` file of the submission.
* Expand grading sheet and bundled submission into normalised folders:
  * All submitted files per user are bundled together if they were not.
  * Links to per-user folder are added to the grading sheet for ease of use.

> :warning: **If you are using MacOS**: 
> Apple's sandboxing, and its implementation in Microsoft Office, 
> might prevent hyperlinks from working. 
> To see if this is the case, 
> move all associated files to 
> `~/Library/Group Containers/<code>.Office/`. 
> Because Microsoft Office has rights for this folder, 
> hyperlinks should then work. 
> (Note that, as of MacOS 12.3 (Monterey), 
> giving Excel `Full disk access` does not solve this issue.)
