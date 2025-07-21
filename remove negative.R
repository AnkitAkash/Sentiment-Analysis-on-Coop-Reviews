# Install and load required packages if not already installed
if (!require("dplyr")) install.packages("dplyr", dependencies=TRUE)
if (!require("readxl")) install.packages("readxl", dependencies=TRUE)
if (!require("openxlsx")) install.packages("openxlsx", dependencies=TRUE)

# Load the required libraries
library(dplyr)
library(readxl)
library(openxlsx)

# Load the strength and weakness files
strength_file <- readxl::read_excel("EMP_word_frequency_strength.xlsx")
weakness_file <- readxl::read_excel("weakness_file.xlsx")

# Extract the list of negative words from the weakness file
negative_words <- weakness_file$weakness

# Ensure the correct column name is used in the strength file
strength_file <- strength_file %>%
  rename(Words_in_Strength = `Words in Strength`)

# Function to check if any word in the text matches a negative word
has_negative_word <- function(text, negative_words) {
  any(grepl(paste(negative_words, collapse = "|"), text, ignore.case = TRUE))
}

# Filter rows in the strength file where the "Words_in_Strength" column has a negative word
strength_file <- strength_file %>%
  filter(!sapply(strength_file$Words_in_Strength, has_negative_word, negative_words))

# Print the modified strength file
print(strength_file)

# Save the modified strength file to a new Excel file
write.xlsx(strength_file, "modifed_strength_file_final.xlsx", rowNames = FALSE)
