get_json <- function(url, tries = 5, timeout = 300) {
  for (i in seq_len(tries)) {
    con <- NULL
    res <- tryCatch({
      h <- curl::new_handle(timeout = timeout)
      con <- curl::curl(url, handle = h)
      on.exit({
        if (!is.null(con) && inherits(con, "connection")) {
          try(close(con), silent = TRUE)
        }
      }, add = TRUE)
      txt <- paste(readLines(con, warn = FALSE), collapse = "\n")
      jsonlite::fromJSON(txt)
    }, error = identity)
    if (!inherits(res, "error")) return(res)
    if (i == tries) stop(res)
    Sys.sleep(1.5^(i - 1))
  }
}

# Robust CSV reader with guarded close
read_csv_url <- function(url, tries = 5, timeout = 300) {
  for (i in seq_len(tries)) {
    con <- NULL
    res <- tryCatch({
      h <- curl::new_handle(timeout = timeout)
      con <- curl::curl(url, handle = h)
      on.exit({ if (!is.null(con) && inherits(con, "connection")) try(close(con), TRUE) }, add = TRUE)
      df <- as.data.frame(readr::read_csv(con, show_col_types = FALSE),
                          stringsAsFactors = FALSE)
      names(df) <- make.names(names(df), unique = TRUE)  # <<< normalize
      df
    }, error = identity)
    if (!inherits(res, "error")) return(res)
    if (i == tries) stop(res)
    Sys.sleep(1.5^(i - 1))
  }
}

urban_rural_sdz_df2 <- read.xlsx(paste0(data_source_root,"geography-data-zone-and-super-data-zone-lookups-v3.xlsx"),
                                 sheet = "SDZ21_Urban_mixed_rural_lookup") %>% 
  rename(`Zone` = SDZ2021_code, `Zone Name` = SDZ2021_name) %>%
  select(Zone, Urban_status, Urban_mixed_rural_status, `Zone Name`)

urban_rural_dz_df2 <- read.xlsx(paste0(data_source_root,"geography-data-zone-and-super-data-zone-lookups-v3.xlsx"),
                                sheet = "DZ21_Urban_mixed_rural_lookup") %>% 
  rename(`Zone` = DZ2021_code, `Zone Name` = DZ2021_name) %>%
  select(Zone, Urban_status, Urban_mixed_rural_status, `Zone Name`)

urban_rural_df2 <- rbind(urban_rural_dz_df2, urban_rural_sdz_df2)

urban_rural_dp <- c("BSDZ", "BSSDZ")

# Importing Flexible Tables
# Create a list of the urls from the Flexible Table Builder
table_list <- c("AGE_BAND_BROAD","AGE_BAND_AGG7_A", "COB_AGG3", "ECONOMIC_ACTIVITY_AGG5",
                "HEALTH_CONDITION_MENTAL_HEALTH", "DISABILITY_DVO", "HEALTH_IN_GENERAL_AGG3", 
                "HH_FAMILY_COMPOSITION_AGG6_PERS", "HH_CAR_VAN_TC2_PERS",
                "HH_DEPENDENT_CHILDREN_IND_PERS", "HH_RENEWABLE_ENERGY_IND_PERS",
                "MAR_CP_STATUS_AGG4", "OCCUPATION_1DIGIT", 
                "TRANSPORT_TO_STUDY_AGG5", "DISTANCE_TO_STUDY_AGG7", "DISTANCE_TO_WORK_AGG8", "TRANSPORT_TO_WORKPLACE_AGG7",
                "IS_CARER_AGG5", "HIGHEST_QUALIFICATION_AGG7", "RELIGION_BELONG_TO_AGG4", "UR_SEX", 
                "SEXUAL_ORIENTATION_DVO_AGG4", 
                "MAIN_LANGUAGE_AGG3")

# Data Zone
list_of_dz_urls <- paste0("https://build.nisra.gov.uk/en/custom/table.csv?d=PEOPLE&v=DZ21&v=", 
                          table_list)

# Super Data Zone
list_of_sdz_urls <- paste0("https://build.nisra.gov.uk/en/custom/table.csv?d=PEOPLE&v=SDZ21&v=",
                           table_list)

# District Electoral Area
list_of_dea_urls <- paste0("https://build.nisra.gov.uk/en/custom/table.csv?d=PEOPLE&v=DEA14&v=",
                           table_list)

# Combined list
list_of_urls <- c(list_of_sdz_urls, list_of_dz_urls, list_of_dea_urls)

# NI Total
list_of_totals <- paste0("https://build.nisra.gov.uk/en/custom/table.csv?d=PEOPLE&v=",
                         table_list)

# ----- Handle MYE CSVs from DataPortal -----
mye_dataportal_urls <- c(
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/MYE01T010/CSV/1.0/en", # dp mye - DEA
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/MYE01T011/CSV/1.0/en", # dp mye - DZ
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/MYE01T012/CSV/1.0/en" # dp mye - SDZ
)

list_of_urls <- c(mye_dataportal_urls, list_of_urls)

# All metadata from data portal read in:
data_portal_all_tables <- get_json(
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadCollection"
)$link$item

# this intersect only returns one result but there are multiple matches
dataset_long_names <- intersect(urban_rural_dp,data_portal_all_tables$extension$matrix)

data_labels <- data.frame(data_portal_all_tables$extension$matrix, data_portal_all_tables$label)

urban_rural_df <- tibble(data_portal_all_tables.extension.matrix = urban_rural_dp)

# joined_data <- left_join(urban_rural_df, data_labels, by = "data_portal_all_tables.extension.matrix")
# joined_data <- joined_data %>%
#   rename("Label" = data_portal_all_tables.label)
joined_data <- c("Benefit Statistics", "Benefit Statistics","Benefit Statistics", "Age (MYE)","Age (MYE)", "Sex (MYE)")

Sys.getenv("http_proxy")
Sys.getenv("https_proxy")
Sys.setenv("http_proxy" = "")
Sys.setenv("https_proxy" = "")

url_list <- c(
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/BSDZ/CSV/1.0/",
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/BSSDZ/CSV/1.0/",
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.ReadDataset/BSDEA/CSV/1.0/",
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T012?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22,%223%22,%224%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D", # MYE: Age
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T010?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22,%223%22,%224%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D",
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T012?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D", # MYE: Sex
  "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T010?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D"
)

list_of_urls <- c(list_of_urls, url_list)

# Create a blank json
final_json <- list("Super Data Zone" = list(),
                   "Data Zone" = list(),
                   "District Electoral Area" = list())

dp_ni_total_list <- list()

date_df <- data.frame(
  Category = character(),
  Year = numeric(),
  stringsAsFactors = FALSE
)

# Loop through each Super Data Zone/Data Zone to add rural/urban/mixed
for (i in 1:length(list_of_urls)) {

  url_name <- list_of_urls[i]
  
  if ((substr(url_name, 1, 13) == "https://ws-da") && 
      !(url_name %in% mye_dataportal_urls)) {
    
    dataset_url <- list_of_urls[i]
    
    csv_data = read_csv_url(dataset_url)
    
    if (dataset_url == "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T012?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22,%223%22,%224%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D") {
      filtered_data <- csv_data[, names(csv_data) %in% c("VALUE", "Broad.age.band..4.cat.", "Year") | 
                                  grepl("dz|sdz|dea", names(csv_data), ignore.case = TRUE)] %>%
        rename("Statistic.Label" = "Broad.age.band..4.cat.", "Zone" = "SDZ2021") %>%
        filter(Year == max(Year, na.rm = TRUE))
      category_name <- "Age (MYE)"
    } else if (dataset_url == "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T010?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22,%223%22,%224%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D") {
      filtered_data <- csv_data[, names(csv_data) %in% c("VALUE", "Broad.age.band..4.cat.", "Year") | 
                                  grepl("dz|sdz|dea", names(csv_data), ignore.case = TRUE)] %>%
        rename("Statistic.Label" = "Broad.age.band..4.cat.", "Zone" = "DEA2014") %>%
        filter(Year == max(Year, na.rm = TRUE))
      category_name <- "Age (MYE)"
    } else if (dataset_url == "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T012?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D") {
      filtered_data <- csv_data[, names(csv_data) %in% c("VALUE", "Sex.Label", "Year") | 
                                  grepl("dz|sdz|dea", names(csv_data), ignore.case = TRUE)] %>%
        rename("Statistic.Label" = "Sex.Label", "Zone" = "SDZ2021")  %>%
        filter(Year == max(Year, na.rm = TRUE))
      category_name <- "Sex (MYE)"
    } else if (dataset_url == "https://ws-data.nisra.gov.uk/public/api.restful/PxStat.Data.Cube_API.PxAPIv1/en/153/PMPE/MYE01T010?query=%7B%22query%22:%5B%7B%22code%22:%22broadage4%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%22All%22%5D%7D%7D,%7B%22code%22:%22Sex%22,%22selection%22:%7B%22filter%22:%22item%22,%22values%22:%5B%221%22,%222%22%5D%7D%7D%5D,%22response%22:%7B%22format%22:%22csv%22,%22pivot%22:null,%22codes%22:true%7D%7D") {
      filtered_data <- csv_data[, names(csv_data) %in% c("VALUE", "Sex.Label", "Year") | 
                                  grepl("dz|sdz|dea", names(csv_data), ignore.case = TRUE)] %>%
        rename("Statistic.Label" = "Sex.Label", "Zone" = "DEA2014")  %>%
        filter(Year == max(Year, na.rm = TRUE))
      category_name <- "Sex (MYE)"
    } else {
      filtered_data <- csv_data[, names(csv_data) %in% c("VALUE", "Statistic.Label", "Year") | 
                                  grepl("dz|sdz|dea", names(csv_data), ignore.case = TRUE)] %>%
        filter(Year == max(Year, na.rm = TRUE))
      matching_cols <- grep("dz|sdz|dea", names(filtered_data), ignore.case = TRUE)
      names(filtered_data)[matching_cols] <- "Zone"
      category_name <- "Benefits Statistics"
    }

    date_df <- rbind(
      date_df,
      data.frame(Category = category_name,
                 Year = max(filtered_data$Year),
                 stringsAsFactors = FALSE))
    
    total_dp_data <- filtered_data %>%
      filter(Zone == "N92000002") %>%
      select(Statistic.Label, VALUE)
    
    total_dp_data$VALUE <- round((total_dp_data$VALUE / sum(total_dp_data$VALUE)) * 100, 1)
    
    dp_ni_total_list[[category_name]] <- setNames(total_dp_data[[2]], total_dp_data[[1]])
    
    filtered_data <- filtered_data %>%
      rename(label_name = "Statistic.Label", "Count" = "VALUE") %>%
      select(-"Year")
    
    cols <- names(filtered_data)
    cols <- cols[cols != "label_name"]
    new_order <- append(cols, "label_name", after = 1)
    filtered_data <- filtered_data[, new_order]
    
    split_df <- split(filtered_data, filtered_data[[1]])
    
  } else if ((substr(url_name, 1, 13) == "https://build") ||
             (url_name %in% mye_dataportal_urls)) {
    
    response <- GET(list_of_urls[i], config(ssl_verifypeer = 0))
    content_text <- content(response, as = "text")
    flexi_data <- read.csv(text = content_text)
    
    #-------------------------- DP MYE data prep ------------------------------
    if (url_name %in% mye_dataportal_urls) {
      
      # create dataframe with columns: Zone, label_name, Count - to match existing logic
      if (grepl("MYE01T010", url_name)) { # DEA Data Portal MYE
        # filter for total population counts and latest year
        flexi_data <- flexi_data %>% 
          filter(broadage4 == "All" & Sex == "All" & Year == max(Year, na.rm = TRUE))
        # DEA: keep col 5,6,12 for consistency - DEA2014, District Electoral Area, value
        flexi_data <- flexi_data[, c(5, 6, 12)]
        # rename columns for later
        colnames(flexi_data)[c(1, 2, 3)] <- c("District Electoral Area 2014 Code", 
                                              "District Electoral Area 2014 Label",
                                              "Count")
      } else if (grepl("MYE01T011", url_name)) { # DZ Data Portal MYE
        # filter for latest year
        flexi_data <- flexi_data %>% 
          filter(Year == max(Year, na.rm = TRUE))
        # DZ: keep col 5,6,7 for consistency - DZ2021, Data Zones, Value
        flexi_data <- flexi_data[, c(5, 6, 8)] 
        # rename columns for later
        colnames(flexi_data)[c(1, 2, 3)] <- c("Census 2021 Data Zone Code", 
                                              "Census 2021 Data Zone Label",
                                              "Count")
      } else if (grepl("MYE01T012", url_name)) { # SDZ Data Portal MYE
        # filter for total population counts and latest year
        flexi_data <- flexi_data %>% 
          filter(broadage4 == "All" & Sex == "All" & Year == max(Year, na.rm = TRUE))
        # SDZ: keep col 5,6,12 for consistency - SDZ2021, Super Data Zones, Value
        flexi_data <- flexi_data[, c(5, 6, 12)]
        # ensure column names are correct
        colnames(flexi_data)[c(1, 2, 3)] <- c("Census 2021 Super Data Zone Code", 
                                              "Census 2021 Super Data Zone Label",
                                              "Count")
      }
    }
    #--------------------------------------------------------------------------
    
    if (ncol(flexi_data) > 3) {
      flexi_data <- flexi_data %>% select(1, 4, 5)
    }
    
    col_name <- tolower(colnames(flexi_data)[2])
    
    if (substr(col_name, 1, 3) == "age") {
      category_name <- gsub("Age\\.\\.\\.(\\d+).*", "Age (\\1 Categories)", colnames(flexi_data)[2])  
    } else {
      category_name <- gsub("\\.{1,2}", " ", sub("\\.\\.\\..*$", "", colnames(flexi_data)[2]))    
    }
    
    split_df <- split(flexi_data, flexi_data[[1]])

    date_df <- rbind(
      date_df,
      data.frame(Category = category_name, 
                 Year = 2021, 
                 stringsAsFactors = FALSE))
    
  } else {
    NA
  }
  
  for (zone in names(split_df)) {
    
    if (substr(zone, 1, 3) == "N20") {
      if (is.null(final_json$`Data Zone`[[zone]])) {
        final_json$`Data Zone`[[zone]] <- list()
      }
      
      final_json$`Data Zone`[[zone]][[category_name]] <- setNames(split_df[[zone]]$Count, split_df[[zone]][[2]])
      
    } else if (substr(zone, 1, 3) == "N21") {
      if (is.null(final_json$`Super Data Zone`[[zone]])) {
        final_json$`Super Data Zone`[[zone]] <- list()
      }
      
      final_json$`Super Data Zone`[[zone]][[category_name]] <- setNames(split_df[[zone]]$Count, split_df[[zone]][[2]])
    } else if (substr(zone, 1, 3) == "N10") {
      if (is.null(final_json$`District Electoral Area`[[zone]])) {
        final_json$`District Electoral Area`[[zone]] <- list()
      }
      
      final_json$`District Electoral Area`[[zone]][[category_name]] <- setNames(split_df[[zone]]$Count, split_df[[zone]][[2]])
    } else {
      NA
    }
  }
}

date_df <- unique(date_df)

date_df <- date_df[!grepl("Data Zone|District Electoral Area", date_df$Category), ]

# Function to add Year to each category
add_year_to_categories <- function(section) {
  for (cat_name in names(section)) {
    year_value <- date_df$Year[date_df$Category == cat_name]
    if (length(year_value) == 1) {
      section[[cat_name]]$Year <- year_value
    }
  }
  return(section)
}

# Apply to each top-level section except "Year"
for (key in names(final_json)) {
  if (key != "Year") {
    final_json[[key]] <- add_year_to_categories(final_json[[key]])
  }
}

final_json$Year <- setNames(date_df$Year, date_df$Category)

zones <- c("Super Data Zone", "Data Zone", "District Electoral Area")

for (zone in zones) {
  final_json[[zone]] <- lapply(final_json[[zone]], convert_to_named_list)
}

# Loop through each Super Data Zone/Data Zone to add rural/urban/mixed
for (i in seq_len(nrow(urban_rural_df2))) {
  
  zone <- urban_rural_df2$`Zone`[i]
  status <- urban_rural_df2$Urban_mixed_rural_status[i]
  
  if (substr(zone, 1, 3) == "N20") {
    if (!is.null(final_json$`Data Zone`[[zone]])) {
      # Extract existing data
      existing_data <- final_json$`Data Zone`[[zone]]
      
      # Rebuild the list with status first
      final_json$`Data Zone`[[zone]] <- c(
        list(Urban_mixed_rural_status = status),
        existing_data
      )
    }
  } else if (substr(zone, 1, 3) == "N21") {
    if (!is.null(final_json$`Super Data Zone`[[zone]])) {
      # Extract existing data
      existing_data <- final_json$`Super Data Zone`[[zone]]
      
      # Rebuild the list with status first
      final_json$`Super Data Zone`[[zone]] <- c(
        list(Urban_mixed_rural_status = status),
        existing_data
      )
    }
  } else {
    NA
  }
}

# Get lgd for each SDZ
# Define the URL and temporary file path
#lgd_sdz_df <-read.csv("T:/General TL/Resources/cpdjul2024/CPDJul2024csv/CPDJul2024csv/CPD_LIGHT.csv", header = TRUE, sep = ",")

# update link to point to data_source_root (MG)
lgd_sdz_df <- read.csv(
  file.path(data_source_root, "CPD_LIGHT.csv"),
  header = TRUE,
  sep = ","
)

lgd_sdz_df_clean <- lgd_sdz_df %>%
  select(5,6,9,10,19,20,21,22) %>%
  drop_na()

# Convert mapping into a named vector
zone_to_lgd <- setNames(lgd_sdz_df_clean$LGD2014NAME, lgd_sdz_df_clean$SDZ2021)

# Loop through each zone in final_json
for (zone in names(final_json$`Super Data Zone`)) {
  
  # Get lgd value using the mapping
  lgd_value <- zone_to_lgd[[zone]]
  
  if (!is.null(lgd_value)) {
    
    existing_data <- final_json$`Super Data Zone`[[zone]]
    
    # Rebuild list with LGD
    final_json$`Super Data Zone`[[zone]] <- c(
      list(Urban_mixed_rural_status = existing_data$Urban_mixed_rural_status),
      list(LGD = lgd_value),
      existing_data[!names(existing_data) %in% c("Urban_mixed_rural_status")]
    )
  }
}

# Convert mapping into a named vector
zone_to_lgd_dz <- setNames(lgd_sdz_df_clean$LGD2014NAME, lgd_sdz_df_clean$DZ2021)

# Loop through each zone in final_json
for (zone in names(final_json$`Data Zone`)) {
  
  # Get lgd value using the mapping
  lgd_value <- zone_to_lgd_dz[[zone]]
  
  if (!is.null(lgd_value)) {
    
    existing_data <- final_json$`Data Zone`[[zone]]
    
    # Rebuild list with LGD
    final_json$`Data Zone`[[zone]] <- c(
      list(Urban_mixed_rural_status = existing_data$Urban_mixed_rural_status),
      list(LGD = lgd_value),
      existing_data[!names(existing_data) %in% c("Urban_mixed_rural_status")]
    )
  }
}

# Convert mapping into a named vector
zone_to_lgd_dea <- setNames(lgd_sdz_df_clean$LGD2014NAME, lgd_sdz_df_clean$DEA2014)

# Loop through each zone in final_json
for (zone in names(final_json$`District Electoral Area`)) {
  
  # Get lgd value using the mapping
  lgd_value <- zone_to_lgd_dea[[zone]]
  
  if (!is.null(lgd_value)) {
    
    existing_data <- final_json$`District Electoral Area`[[zone]]
    
    # Rebuild list with LGD
    final_json$`District Electoral Area`[[zone]] <- c(
      list(Urban_mixed_rural_status = existing_data$Urban_mixed_rural_status),
      list(LGD = lgd_value),
      existing_data[!names(existing_data) %in% c("Urban_mixed_rural_status")]
    )
  }
}

# Initialize NI Total list
ni_total_list <- list()

# Loop through each NI Total URL
for (i in seq_along(list_of_totals)) {

  # Load NI-level data
  response <- GET(list_of_totals[i], config(ssl_verifypeer = 0))
  content_text <- content(response, as = "text")
  flexi_data <- read.csv(text = content_text, stringsAsFactors = FALSE)
  
  # Select only the first two columns if more are present
  if (ncol(flexi_data) > 2) {
    flexi_data <- flexi_data[, 2:3]
  }
  
  col_name <- names(flexi_data)[2]
  
  flexi_data <- flexi_data %>%
    mutate(!!col_name := round((.data[[col_name]] / sum(.data[[col_name]])) * 100, 1))
  
  column_name <- tolower(colnames(flexi_data)[1])
  
  if (substr(column_name, 1, 3) == "age") {
    category_name <- gsub("Age\\.\\.\\.(\\d+).*", "Age (\\1 Categories)", colnames(flexi_data)[1])  
  } else {
    category_name <- gsub("\\.{1,2}", " ", sub("\\.\\.\\..*$", "", colnames(flexi_data)[1]))    
  }
  
  # Ensure the first column is character (attribute labels)
  flexi_data[[1]] <- as.character(flexi_data[[1]])
  
  # Add named vector directly to the list
  ni_total_list[[category_name]] <- setNames(flexi_data[[2]], flexi_data[[1]])
}

final_ni_total_list <- c(ni_total_list, dp_ni_total_list)

final_json <- c(list("NI Total" = final_ni_total_list), final_json)

# Create lookup table for category codes in the urls
ft_ni_total_list <- c(ni_total_list)
nested_list_names <- names(ft_ni_total_list)
nested_list_names <- unlist(nested_list_names)
further_breakdown_df <- table_list

category_lookup <- data.frame(nested_list_names, further_breakdown_df)
lookup_data <- category_lookup[, c("further_breakdown_df", "nested_list_names")]
lookup_data$Source <- "Flexible Table Builder"
  
data_portal_lookup <- data.frame("further_breakdown_df" = c("Benefits", "MYE01T012", "MYE01T012"),
  "nested_list_names" = c("Benefits Statistics", "Age (MYE)", "Sex (MYE)"),
  "Source" = c("Data Portal", "Data Portal", "Data Portal"))

lookup_data <- rbind(lookup_data, data_portal_lookup)


# Attach Year from final_json to lookup_data

year_vec <- final_json[["Year"]]
year_df <- data.frame(
  nested_list_names = names(year_vec),
  Year = as.numeric(unname(year_vec)),
  stringsAsFactors = FALSE
)

# Preserve order and rows
lookup_data <- merge(lookup_data, year_df, by = "nested_list_names", all.x = TRUE, sort = FALSE)

# Write json
write_json(lookup_data, "category_lookup.json", pretty = TRUE)
