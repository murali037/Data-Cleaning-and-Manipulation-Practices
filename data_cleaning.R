# 
install.packages("pdftools")
install.packages("tidyverse")
install.packages("readr")

library(pdftools)
library(tidyverse)
library(readr)
library(stringr)
library(xlsx)
library(glue)

extract_text <- pdf_text("Bryan_MS_thesis_final.pdf") %>% readr::read_lines()

extract_text[1703:1737]

table <- extract_text[2000:2200]

table_1 <- table %>% str_squish() #%>% strsplit(split = " ")


page_51 <- extract_text[1703:1737]


page_51 <- page_51 %>% str_squish()



write.xlsx(page_51, "page_51.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)

head(table_1)


write.xlsx(table_1, "table.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)


############################### table 2.4  #############################################################
#works without headers

extract_text[1027]

table_2.4 <- extract_text[1029:1084]

table_2.4 <- table_2.4 %>% strsplit(split=" ")

table_2.4

species_cat <- extract_text[1027] %>% unlist()
# species_cat[c(2,3,4,5,6,7,8,9,10,11)] <- c("C. pachyderma_Li/Ca", "C. pachyderma_Mg/Li",
#                                            "U. peregrina_Li/Ca", "U. peregrina_Mg/Li",
#                                            "P. ariminensis_Li/Ca", "P. ariminensis_Mg/Li",
#                                            "P. foveolata_Li/Ca", "P. foveolata_Mg/Li",
#                                            "H. elegans_Li/Ca", "H. elegans_Mg/Li")
species_cat


#table_2.4.df <- plyr::ldply(table_2.4)

write.xlsx(table_2.4, "v2_table_2.4.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)


#2.3, 3.2


############################### table 2.3  #############################################################

extract_text[553:608]

table_2.3 <- extract_text[545:608]

table_2.3 

table_2.3 <- table_2.3 #%>% strsplit(split = " ")

table_2.3

#table_2.3.df <- plyr::ldply(table_2.3)

write.xlsx(table_2.3, "table_2.3.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)


############################### table 3.2  #############################################################

extract_text[1700:1850]


#################################################################################################################

extract_text_2 <- pdf_text("Tang2008Geochim.Cosmochim.Acta.pdf") %>% readr::read_lines()

table_2 <- extract_text_2[151:193]

extract_text_2[151:193]

table_2 <- table_2 %>% str_squish()


write.xlsx(table_2, "table_2.xlsx", sheetName = "Sheet1", 
           col.names = TRUE, row.names = TRUE, append = FALSE)



#################################################################################################################





