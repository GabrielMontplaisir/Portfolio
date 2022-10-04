# Portfolio
A script to create Portfolios for a group of students (in the form of Google Slides). When students submit a Google Form, append a new slide with their responses and your comments into their own Portfolio so they can find all the information in one place.

## Installation

1. Create a copy of Portfolio by clicking on the link in the description, or copy the code into your own Sheet.
2. Authenticate the extension by selecting Extensions > Portfolio > Show Sidebar.
3. Import Google Forms by selecting `Import Form` or create a new form via `Create Form`.
4. In the last column of a Form Response Tab, add a column with the header `Comments`. Add comments for student responses.
5. Via the Hamburger Icon, select the `Comments Column` header from one of the Forms response tabs. This was done to support the different languages for Google Forms.

## Usage

Once you've set everything up from the [Installation](#installation) section, select one of the response tabs and select the `Export to Portfolio` option. The first time you execute this script, it will longer than usual. This is because the script is generating portfolios for all your students and placing them in a dedicated folder in your Google Drive (to keep everything organized). The portfolios are automatically shared with the students with Editor access. If at any point, the student deletes their portfolio, you delete the portfolio, or the link is deleted, the script will automatically create a new Portfolio for that student.

When you click `Export to Portfolio`, a slide is appended to each student's portfolio with their responses, the teacher comments and a section for notes. At this time, you have to select `Export to Portfolio` for each tab.

**Note: The tab you wish to export must be linked to a Google Form. If it isn't, the script will return an error. Either import the form from the Sidebar, or change the destination inside the Google Form itself.**

## Coming Soon

- Export ALL Google Form Tabs to Portfolios.
- Add an option to create Portfolios for all students
- Allow the user to select the Template Slide of their choosing
- Save and find the Portfolios even if the link was erased in the Portfolio tab.
- Save, find and modify the appropriate Slide in the Portfolio for the appropriate response (unsure if this is possible). 
- Fix bugs & Optimize code where possible
