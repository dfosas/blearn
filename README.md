# blearn

[![Code style: black](
https://img.shields.io/badge/code%20style-black-000000.svg)](
https://github.com/psf/black)

Utilities to work with BlackBoard Learn in offline mode (i.e., *not* with its API).

Package under construction:
* No that much documentation yet.
  * Best entry points are CLI tools starting with `blearn-` (they have docs).
* No tests yet (it would require widespread availability of test courses
  for known Learn versions to make them truly reproducible).
* Things can easily break with Learn updates.
* Offline marking of non-anonymous submissions still needs some love to handle
  reporting of some cases (e.g., what happens with repeated submissions).

# Installation
Regular installation for python packages only available at git apply.
The package is being developed with `flit`, so `flit install --symlink` works great.
Since the package includes web automation routines with `playwright`,
it may be necessary to also install its tools with `playwright install`
*after* this package is installed to make browsers available.

# Functionality
* Parse offline grading sheet.
* Parse bundled submission in zip file or extracted zip file:
  * Metadata based on file names.
  * Metadata from the content in the log `.txt` file of the submission.
* Expand grading sheet and bundled submission into normalised folders:
  * All submitted files per user are bundled together if they were not.
  * Links to per-user folder are added to the grading sheet for ease of use.
* Support for anonymous marking that is not made available by Learn (!).

> :warning: **If you are using macOS**: 
> Apple's sandboxing, and its implementation in Microsoft Office, 
> might prevent hyperlinks from working. 
> To see if this is the case, 
> move all associated files to 
> `~/Library/Group Containers/<code>.Office/`. 
> Because Microsoft Office has rights for this folder, 
> hyperlinks should then work. 
> (Note that, as of macOS 12.3 (Monterey), 
> giving Excel `Full disk access` does not solve this issue.)

> :warning: **If you are using folders synced to the cloud**:
> In cases like OneDrive on Windows and Excel,
> the OS might redirect paths to files synced to the cloud by a URL.
> In that case, it might be preferable to work on a fully offline location.
