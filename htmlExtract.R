#################
# Title: HTML Extract
# Author: Andrew DiLernia
# Date: 05/27/2020
# Purpose: Extract and export data from HTML tables to Excel
#################

library(tidyverse)
library(rvest)

# f <- list.files("Eval")[42]
# aggResEx <- readRDS("exampleDF.rds")
# which(list.files("Eval") %in% errorFiles)

errorFiles <- c()

completeRes <- map_dfr(.x = list.files("Eval"), .f = function(f) {
  aggRes <- data.frame()
  responseData <- data.frame()
  
try({
# Reading in HTML file
htmlData <- xml2::read_html(paste0("Eval/", f))

# Table counter off-set for oddly structured files
counter <- 0

# Questions 1 through 3
tabData123 <- htmlData %>% html_nodes(xpath='/html/body/form/table[3]') %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(nrow(tabData123) == 13) {
  # Student responses
  for(q in 1:3) {
    responseData <- rbind(responseData, setNames(c(tabData123[6, 5], tabData123[9 + q, 2:14]), 
                                                 c("Faculty", substr(tabData123[9 + q, 1], 1, 1), tabData123[9, 14:25])) %>% as.data.frame() %>% 
                            rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
                                   Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
  }
} else if(nrow(tabData123) == 11) {
  
  counter <- 2
  
  for(q in 1:3) {
  # Question 2
  tabData123 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 2+q, ']')) %>%
    html_table(fill = TRUE) %>% as.data.frame()
  
  responseData <- rbind(responseData, setNames(c(tabData123[6, 5], tabData123[10, 2:14]), 
                                               c("Faculty", substr(tabData123[10, 1], 1, 1), tabData123[9, 14:25])) %>% as.data.frame() %>% 
                          rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
                                 Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
  }
  
} else {stop("Issue with scraping questions 1 through 3.")}

# Extracting data for Q4
tabData4 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 4 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

temp <- setNames(c(tabData4[6, 5], tabData4[10, 2:14]), c("Faculty", "Q", tabData4[9, 14:25])) %>%
  as.data.frame()

if("E" %in% colnames(temp)) {
responseData <- rbind(responseData, temp %>% 
                        rename(StronglyAgree = E, Agree = `E.1`, SomewhatAgree = VG, SomewhatDisagree = S,
                               Disagree = P, StronglyDisagree = VP, Med = `Med.`))
} else {
  responseData <- rbind(responseData, temp %>% as.data.frame() %>% 
                          rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
                                 Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
}

# Extracting data for Q5
tabData <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 5 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(tabData[10, 1] == "Q3" & tabData[10, 2] != "The instructor facilitated learning activities (e.g., discussion, assignments, readings) that deepened my understanding of health inequalities.") {
  tabData <- htmlData %>% html_nodes(xpath='/html/body/form/table[7]') %>%
    html_table(fill = TRUE) %>% as.data.frame()
} else if(tabData[10, 1] == "Q3") {
  tabData[10, 1] <- "Q5"
}

temp <- setNames(c(tabData[6, 5], tabData[10, 2:14]), c("Faculty", "Q", tabData[9, 14:25]))

if("SA" %in% names(temp)) {
responseData <- rbind(responseData, temp %>%
                        as.data.frame() %>% 
  rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
         Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
} else {
  responseData <- rbind(responseData, temp %>% 
                          as.data.frame() %>% 
                          rename(StronglyAgree = E, Agree = `E.1`, SomewhatAgree = VG, SomewhatDisagree = S,
                                 Disagree = P, StronglyDisagree = VP, Med = `Med.`))
}

# Questions 6 through 7
tabData67 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 6 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(ncol(tabData67) == 2) {
  tabData67 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 6, ']')) %>%
    html_table(fill = TRUE) %>% as.data.frame()
  
  for(q in 1:2) {
    responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData67[9 + q, 2:14]), 
                                                 c("Faculty", substr(tabData67[9 + q, 1], 1, 1), tabData67[9, 14:25])) %>% as.data.frame() %>% 
                            rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
                                   Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
  }
  
} else {
# Student responses
for(q in 1:2) {
  responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData67[9 + q, 2:14]), 
                        c("Faculty", substr(tabData67[9 + q, 1], 1, 1), tabData67[9, 14:25])) %>% as.data.frame() %>% 
                          rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, SomewhatDisagree = SD,
                                 Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
}
}

# Extracting data for Q8
tabData8 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 7 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(ncol(tabData8) == 2) {
  tabData8 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 16, ']')) %>%
    html_table(fill = TRUE) %>% as.data.frame()
  responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData8[10, 2:14]), c("Faculty", "Q", tabData8[9, 14:25])) %>% 
                          as.data.frame() %>% 
                          rename(StronglyAgree = E, Agree = `E.1`, SomewhatAgree = VG, SomewhatDisagree = S,
                                 Disagree = P, StronglyDisagree = VP, Med = `Med.`))
  
} else {
responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData8[10, 2:14]), c("Faculty", "Q", tabData8[9, 14:25])) %>% 
                        as.data.frame() %>% 
                        rename(StronglyAgree = E, Agree = `E.1`, SomewhatAgree = VG, SomewhatDisagree = S,
                               Disagree = P, StronglyDisagree = VP, Med = `Med.`))
}

# Extracting data for Q9
tabData9 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 8 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(ncol(tabData9) > 2) {
  responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData9[10, 2:14]), c("Faculty", "Q", tabData9[9, 14:25])) %>% 
                          as.data.frame() %>% rename(StronglyAgree = SA, Agree = A, SomewhatAgree = `SA.1`, 
                                                     SomewhatDisagree = SD, Disagree = D, StronglyDisagree = `SD.1`, Med = `Med.`))
}

# Questions 12 - 15
tabData1215 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 13 + counter, ']')) %>%
  html_table(fill = TRUE) %>% as.data.frame()

if(nrow(tabData1215) == 0) {
  tabData1215 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 11, ']')) %>%
    html_table(fill = TRUE) %>% as.data.frame()
}

if(nrow(tabData1215) == 2) {
  tabData1215 <- htmlData %>% html_nodes(xpath=paste0('/html/body/form/table[', 12 + counter, ']')) %>%
    html_table(fill = TRUE) %>% as.data.frame()
}

if(ncol(tabData1215) > 2) {
responseData <- responseData %>% mutate(Very = NA, Moderately = NA, Somewhat = NA, NotAtAll = NA)

responseData <- rbind(responseData, setNames(c(tabData[6, 5], tabData1215[15:18, 2:12]), c("Faculty", "Q", tabData1215[14, 12:21])) %>% 
                        as.data.frame() %>% rename(Very = V, Moderately = M, Somewhat = S, 
                                                   NotAtAll = NAA, Med = `Med.`) %>% 
                        mutate(StronglyAgree = NA, Agree = NA, SomewhatAgree = NA, SomewhatDisagree = NA, 
                               Disagree = NA, StronglyDisagree = NA) %>% 
select(Faculty, Q, StronglyAgree, Agree, SomewhatAgree, SomewhatDisagree, Disagree, 
       StronglyDisagree, N, Mean, GrpMed, Med, Mode, StdDev, Very, Moderately, Somewhat, NotAtAll))
}

if(tabData[6, 5] == "7210-UMNTC - 102") {
  responseData$Faculty <- as.data.frame(html_table(html_nodes(htmlData, xpath='/html/body/form/table[2]'),
                                                   fill = TRUE))[4, 2]
} else {
responseData$Faculty <- tabData[6, 5]
}

# Course and department info
courseinfo <- htmlData %>% html_nodes(xpath='/html/body/form/table[2]') %>%
  html_table(fill = TRUE) %>% as.data.frame()

courseCols <- setNames(c(courseinfo[3, 2], courseinfo[3, 4], gsub(map(str_split(courseinfo[4, 4], pattern = "[(]"), 2), 
                         pattern = "[%)]", replacement = "")),
                       c(gsub(courseinfo[3, 1], pattern = "[:]", replacement = ""), 
                         gsub(courseinfo[3, 3], pattern = "[:]", replacement = ""), "PercentResponse"))

# Combining
aggRes <- t(courseCols) %>% cbind(responseData)

# Extracting term
term <- htmlData %>% html_nodes(xpath='/html/body/form/table[1]') %>%
  html_table(fill = TRUE) %>% map(1) %>% str_split(pattern = "_")

if(length(term[[1]]) > 1) {
aggRes$Term <- term %>% map_chr(2)
} else {
  aggRes$Term <- term %>% unlist()
}

aggRes <- aggRes %>% select(Term, Department, Course, Faculty, everything())

return(aggRes)
}, silent = TRUE)
  
  if(nrow(aggRes) == 0) {
    errorFiles <<- c(errorFiles, f)
    return(data.frame())
  }
})

# Cleaning up final table for small typos
completeRes <- completeRes %>% mutate_all(.funs = trimws) %>% 
  mutate(PercentResponse = as.numeric(PercentResponse),
    Term = case_when(
    Term == "Fall2018" ~ "Fall 2018",
    Term == "EpiCH" ~ "Fall 2018",
    Term == "Summer2019" ~ "Summer 2019",
    Term == "003" ~ "Summer 2019",
    str_detect(Term, pattern = "7200-001") ~ "Summer 2019",
    str_detect(Term, pattern = "7250") ~ "Summer 2019",
    Term == "Spring2019" ~ "Spring 2019",
    Term == "002" ~ "Spring 2019",
    TRUE ~ Term
  ), Department = case_when(
    Department == "SPH EpiCH" ~ "EpiCH",
    Department == "SPH HPM" ~ "HPM",
    Department == "SPH Bio" ~ "Biostat",
    Department == "SPH EnHS" ~ "EHS",
    Department == "SPH Inst" ~ "Inst",
    Department == "SPH PHP" ~ "PHP"))

# Making last columns numeric
completeRes[, 7:18] <- lapply(completeRes[, 7:18], FUN = function(x) {
  as.numeric(sapply(x, FUN = function(s) {
    return(max(as.numeric(unlist(str_split(s, pattern = ",")))))
  }))
  })

# Formatting appropriately
completeRes <- completeRes %>% mutate(Year = map_chr(str_split(Term, pattern = " "), 2),
                                        Term = map_chr(str_split(Term, pattern = " "), 1),
                                        Division = Department, Question = Q,
                                        CourseNumber = map_chr(str_split(Course, pattern = "-"), 1),
                                        Title = map_chr(str_split(Course, pattern = "-"), 3)) %>% 
  select(-Course, -Department, -Q) %>% rename(Course = CourseNumber) %>% 
  select(Year, Term, Title, Course, Division, everything(), PercentResponse, Question)

completeRes2 <- completeRes %>% filter(!is.na(StronglyAgree)) %>% select(-c(Very:NotAtAll))

completeResEpiCH <- completeRes %>% filter(!is.na(Very)) %>% select(-c(StronglyAgree:StronglyDisagree))

# Exporting to Excel
writexl::write_xlsx(x = completeRes2, path = "surveySummary.xlsx")
writexl::write_xlsx(x = completeResEpiCH, path = "surveySummary_EpiCH_extra.xlsx")
