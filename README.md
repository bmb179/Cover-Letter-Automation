# Rapidly Generate Cover Letters with R and VBA

Everyone's least favorite vestige of pre-internet hiring that just will not die. 

In my experience, these seem to be just a formality that either don't get read, get skimmed for 10 seconds by a recruiter, or are parsed by an ATS software. When I've been interviewed for roles that I submitted a cover letter for, the interviewer almost never seems to have a recollection of points from the cover letter that aren't listed on my resume. As a result, I usually don't submit one unless it's required. In my view, a bad cover letter can hurt your candidacy significantly more than including a good one can help it.

I created this tool for the cases of employers that still require them.

After you have a few pending applications that require a cover letter and you've saved the variables associated with each role in a CSV file, you can generate multiple cover letters personalized to each of the postings using the R and VBA scripts included in this repo using a template saved as a text file.

The R script will output the personalized cover letters as text files. When I was actively using this in 2023, I was opening them all individually with Word and running a VBA macro that typesets the cover letter and saves it as a PDF. I had to recreate the VBA script for the purpose of this repo since it was only saved locally on an old machine that I no longer have. The goal was to eventually replace this last part in the process with R code using the officeR package, but I never got around to it at the time.

Today I'm publishing the code in case someone finds any value in this or wants to finish the project.
