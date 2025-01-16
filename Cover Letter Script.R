################################################
# A SIMPLE COVER LETTER GENERATOR WRITTEN IN R #
################################################
# My Process:
# - Build a CSV with the variables below.
# - Make a text file template with placeholders for the variables.
# - Run the code to write new cover letter text files.
# - Open them individually with Word and run a VBA macro that typesets the cover letter and saves it as a PDF.
# The goal was to eventually replace this last part in the process with R code using the officeR package, but I never got around to it.

setwd('placeholder_dir')
jobs <- read.csv('contacts.csv')
date <- Sys.Date()

setwd('final_dir')

for (i in 1:nrow(jobs)) {
  biz.name <- jobs$Business.Name[i]
  position <- jobs$Position[i]
  hrg.mgr.1 <- jobs$Hiring.Manager.1[i]
  hrg.mgr.2 <- jobs$Hiring.Manager.2[i]
  hrg.mgr.title <- jobs$Hiring.Manager.Title[i]
  biz.address <- jobs$Business.Address[i]
  city.state.zip <- jobs$Business.City.State.Zip[i]

  CL.Template <- readLines('CL_Template_Sample.txt')
  
  CL.Template <- gsub('\\[Date\\]', date, CL.Template)
  CL.Template <- gsub('\\[Business Name\\]', biz.name, CL.Template)
  CL.Template <- gsub('\\[Position\\]', position, CL.Template)
  CL.Template <- gsub('\\[Hiring Manager 1\\]', hrg.mgr.1, CL.Template) 
  # This is the hiring manager name listed in the letterhead. It would be left blank along with "Hiring Manager Title" if I couldn't find one.
  CL.Template <- gsub('\\[Hiring Manager 2\\]', hrg.mgr.2, CL.Template)
  # This is the hiring manager name listed in the greeting. 
  # If I couldn't find a name, I would use the placeholder "Hiring Manager" in my CSV so the final product would be phrased as "Dear Hiring Manager,"
  CL.Template <- gsub('\\[Hiring Manager Title\\]', hrg.mgr.title, CL.Template)
  CL.Template <- gsub('\\[Business Address\\]', biz.address, CL.Template)
  CL.Template <- gsub('\\[Business City State Zip\\]', city.state.zip, CL.Template)
  
  writeLines(CL.Template, paste0('Cover Letter ', jobs$Business.Name[i], '.txt'))
}