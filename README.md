# GDocHeaderLinker
This script parses the table of content in a google doc within a given range. It then looks for matching text in the doucment and links them to that given header. It can be easily adjusted using entries of the table of content to define ranges to prase. However but it may not ignore matching headers in the document itself depending on level of indentation.

This script also takes a few minutes to run as it scans the whole looking for a match, for each header parsed from the table of contents. This script is not polished and is a more of a proof of concept. However, I thought someone may find use of it, or use in expanding it, or I may revisit it one day. 

## How to use
 * At the top of a google doc file, go to Extensions > Apps Script. Copy and paste the code there.
 * Set startStopPairs, each pair defines a range between two entries on the table of contents that this program will prase. It'll skip over anything else.
 * Select the "LinkHeaders()" as the function to run and run the "LinkHeaders()" function... that's it.
 * You may want to check the final product in case it added the links in areas you didn't want it.

# Credit
Thanks to https://stackoverflow.com/a/18344552 for the code to parse the Table of Contents within a google doc.
