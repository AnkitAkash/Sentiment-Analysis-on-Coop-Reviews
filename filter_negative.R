# Install and load required packages if not already installed
if (!require("dplyr")) install.packages("dplyr", dependencies=TRUE)
if (!require("readxl")) install.packages("readxl", dependencies=TRUE)
if (!require("openxlsx")) install.packages("openxlsx", dependencies=TRUE)
if (!require("tidytext")) install.packages("tidytext", dependencies=TRUE)

# Load the required libraries
library(dplyr)
library(readxl)
library(openxlsx)
library(tidytext)

# Load your strength file
strength_file <- readxl::read_excel("EMP_word_frequency_strength.xlsx")

# Ensure the correct column name is used in the strength file
strength_file <- strength_file %>%
  rename(Words_in_Strength = `Words in Strength`)

# Tokenize the text in the "Words_in_Strength" column
strength_tokens <- strength_file %>%
  unnest_tokens(word, Words_in_Strength)

# Get the sentiment of each word
sentiments <- get_sentiments("bing")
strength_tokens <- inner_join(strength_tokens, sentiments, by = "word")

# Identify row numbers with negative sentiment
negative_rows <- strength_tokens %>%
  filter(sentiment == "negative") %>%
  select(row_number = row_number())

# Remove rows with negative sentiment from the original strength file
strength_file_filtered <- strength_file %>%
  anti_join(negative_rows, by = c("row_number" = "row_number"))

# Print the modified strength file
print(strength_file_filtered)

# Save the modified strength file to a new Excel file
write.xlsx(strength_file_filtered, "modified_strength_file.xlsx", rowNames = FALSE)
