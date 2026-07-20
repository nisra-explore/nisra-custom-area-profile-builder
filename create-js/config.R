## config ##

#### install packages ####

suppressWarnings(if(!require(pacman)) install.packages("pacman"))
library(pacman)
p_load("dplyr",
       "openxlsx", 
       "readxl",
       "rjson",
       "jsonlite",
       "janitor",
       "gtools",
       "httr",
       "stringr",
       "readr",
       "tidyr",
       "tibble",
       "purrr")

#### set-up folders & file names ####
data_source_root <- "create-js/inputs/"


# Turn easier read URL queries to valid URLs
transform_URL <- function(URL) {

  URL %>%
    gsub(" ", "%20", .) %>%
    gsub('"', "%22", .) %>%
    gsub("\\{", "%7B", .) %>%
    gsub("\\}", "%7D", .) %>%
    gsub("\\[", "%5B", .) %>%
    gsub("\\]", "%5D", .)

}

convert_named_vectors <- function(x) {
  if (is.list(x)) {
    lapply(x, convert_named_vectors)
  } else if (!is.null(names(x))) {
    as.list(x)
  } else {
    x
  }
}

# Convert all named vectors to named lists
convert_to_named_list <- function(x) {
  lapply(x, function(category) {
    as.list(category)
  })
}
