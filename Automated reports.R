rm(list = c()) 
options(repos = c(CRAN = "http://cran.r-project.org"))


# Install packages (ignoring if already installed)
required_packages <- c("readxl", "MASS", "ggplot2", "officer", "flextable", "tidyr", "dplyr", "magrittr", "writexl")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)

# Load the necessary libraries into the R session
library(ggplot2)  
library(readxl)   
library(MASS)     
library(officer)  
library(flextable) 
library(tidyr)
library(dplyr)
library(magrittr)
library(writexl) 

# Loads and transforms original data structure
load_and_transform_data <- function() {
    file_path <- file.choose()
    data <- read_excel(file_path, sheet = 1)
    
    # Adds additional columns of IDProvider, SystemicABX, SpectrumScore, DOT, UniqueMDEncounter
    data <- data %>%
      mutate(IDProvider = ifelse(OrderingPhysicianName %in% c(
        "Chris Clinichill", "Chris Oncoberg", "Pat Stethoway", "Pat Pharmaman", 
        "Robin Infectostone", "Robin Nephrohill", "Pat Medistone", "Jordan Theraman", 
        "Alex Scriptlytic", "Casey Therascope", "Alex Gastroman", "Alex Infectostone", 
        "Devon Mediscope", "Lee Nephroline", "Drew Medinova", "Taylor Oncostone", 
        "Drew Scriptwell", "Robin Pharmastone", "Morgan Scriptberg"), 1, 0) # NOTE: Replace the OrderPhysicianNames with your institution-specific prescriber names of interest
,
        SystemicABX = ifelse(Route %in% c('intravenous','oral','g-tube','feeding tube','intramuscular',
        'nasogastric','nasogastric tube', 'NG/OG tube'), 1, 0), # NOTE: Replace the route names with your institution-specific route names
        SpectrumScore = case_when(
          grepl("eravacycline", DrugName, ignore.case = TRUE) ~ 22,
          grepl("omadacycline", DrugName, ignore.case = TRUE) ~ 22,
          grepl("tigecycline", DrugName, ignore.case = TRUE) ~ 22,
          grepl("cefiderocol", DrugName, ignore.case = TRUE) ~ 13,
          grepl("levofloxacin", DrugName, ignore.case = TRUE) ~ 17,
          grepl("imipenem", DrugName, ignore.case = TRUE) & grepl("relebactam", DrugName, ignore.case = TRUE) ~ 15,
          grepl("colistimethate", DrugName, ignore.case = TRUE) | grepl("colistin", DrugName, ignore.case = TRUE) ~ 14,
          grepl("meropenem", DrugName, ignore.case = TRUE) & grepl("vaborbactam", DrugName, ignore.case = TRUE) ~ 14,
          grepl("delafloxacin", DrugName, ignore.case = TRUE) ~ 14,
          grepl("polymyxin", DrugName, ignore.case = TRUE) ~ 13,
          grepl("sulfamethoxazole", DrugName, ignore.case = TRUE) & grepl("trimethoprim", DrugName, ignore.case = TRUE) ~ 13,
          grepl("tobramycin", DrugName, ignore.case = TRUE) ~ 13,
          grepl("fosfomycin", DrugName, ignore.case = TRUE) ~ 13,
          grepl("imipenem", DrugName, ignore.case = TRUE) & !grepl("relebactam", DrugName, ignore.case = TRUE) ~ 12,
          grepl("meropenem", DrugName, ignore.case = TRUE) & !grepl("vaborbactam", DrugName, ignore.case = TRUE) ~ 12,
          grepl("doripenem", DrugName, ignore.case = TRUE) ~ 12,
          grepl("ciprofloxacin", DrugName, ignore.case = TRUE) ~ 12,
          grepl("moxifloxacin", DrugName, ignore.case = TRUE) ~ 12,
          grepl("amikacin", DrugName, ignore.case = TRUE) ~ 12,
          grepl("nitrofurantoin", DrugName, ignore.case = TRUE) ~ 12,
          grepl("ceftazidime", DrugName, ignore.case = TRUE) & grepl("avibactam", DrugName, ignore.case = TRUE) ~ 11,
          grepl("piperacillin", DrugName, ignore.case = TRUE) & grepl("tazobactam", DrugName, ignore.case = TRUE) ~ 11,
          grepl("gentamicin", DrugName, ignore.case = TRUE) ~ 11,
          grepl("plazomicin", DrugName, ignore.case = TRUE) ~ 11,
          grepl("ticarcillin", DrugName, ignore.case = TRUE) & grepl("clavulanate", DrugName, ignore.case = TRUE) ~ 10,
          grepl("minocycline", DrugName, ignore.case = TRUE) ~ 10,
          grepl("ertapenem", DrugName, ignore.case = TRUE) ~ 9,
          grepl("ceftolozane", DrugName, ignore.case = TRUE) & grepl("tazobactam", DrugName, ignore.case = TRUE) ~ 9,
          grepl("gemifloxacin", DrugName, ignore.case = TRUE) ~ 9,
          grepl("chloramphenicol", DrugName, ignore.case = TRUE) ~ 9,
          grepl("ampicillin", DrugName, ignore.case = TRUE) & grepl("sulbactam", DrugName, ignore.case = TRUE) ~ 9,
          grepl("cefepime", DrugName, ignore.case = TRUE) ~ 8,
          grepl("ceftaroline", DrugName, ignore.case = TRUE) ~ 7,
          grepl("norfloxacin", DrugName, ignore.case = TRUE) ~ 7,
          grepl("ofloxacin", DrugName, ignore.case = TRUE) ~ 7,
          grepl("piperacillin", DrugName, ignore.case = TRUE) & !grepl("tazobactam", DrugName, ignore.case = TRUE) ~ 7,
          grepl("doxycycline", DrugName, ignore.case = TRUE) ~ 7,
          grepl("tetracycline", DrugName, ignore.case = TRUE) ~ 7,
          grepl("amoxicillin", DrugName, ignore.case = TRUE) & grepl("clavul", DrugName, ignore.case = TRUE) ~ 7,
          grepl("lefamulin", DrugName, ignore.case = TRUE) ~ 7,
          grepl("clindamycin", DrugName, ignore.case = TRUE) ~ 6,
          grepl("daptomycin", DrugName, ignore.case = TRUE) ~ 6,
          grepl("cefotaxime", DrugName, ignore.case = TRUE) ~ 6,
          grepl("ceftazidime", DrugName, ignore.case = TRUE) ~ 6,
          grepl("ceftriaxone", DrugName, ignore.case = TRUE) ~ 6,
          grepl("oritavancin", DrugName, ignore.case = TRUE) ~ 6,
          grepl("aztreonam", DrugName, ignore.case = TRUE) ~ 6,
          grepl("azithromycin", DrugName, ignore.case = TRUE) ~ 6,
          grepl("linezolid", DrugName, ignore.case = TRUE) ~ 6,
          grepl("tedizolid", DrugName, ignore.case = TRUE) ~ 6,
          grepl("cefoxitin", DrugName, ignore.case = TRUE) ~ 6,
          grepl("cefotetan", DrugName, ignore.case = TRUE) ~ 6,
          grepl("telavancin", DrugName, ignore.case = TRUE) ~ 5,
          grepl("amoxicillin", DrugName, ignore.case = TRUE) & !grepl("clavul", DrugName, ignore.case = TRUE) ~ 5,
          grepl("ampicillin", DrugName, ignore.case = TRUE) & !grepl("sulbactam", DrugName, ignore.case = TRUE) ~ 5,
          grepl("dalbavancin", DrugName, ignore.case = TRUE) ~ 5,
          grepl("vancomycin", DrugName, ignore.case = TRUE) & !grepl("oral", DrugName, ignore.case = TRUE) ~ 5,
          grepl("clarithromycin", DrugName, ignore.case = TRUE) ~ 5,
          grepl("quinupristin", DrugName, ignore.case = TRUE) & grepl("dalfopristin", DrugName, ignore.case = TRUE) ~ 5,
          grepl("rifampin", DrugName, ignore.case = TRUE) ~ 5,
          grepl("cefpodoxime", DrugName, ignore.case = TRUE) ~ 4,
          grepl("cefdinir", DrugName, ignore.case = TRUE) ~ 4,
          grepl("cefaclor", DrugName, ignore.case = TRUE) ~ 4,
          grepl("cefprozil", DrugName, ignore.case = TRUE) ~ 4,
          grepl("cefuroxime", DrugName, ignore.case = TRUE) ~ 4,
          grepl("cefixime", DrugName, ignore.case = TRUE) ~ 3,
          grepl("cefadroxil", DrugName, ignore.case = TRUE) ~ 3,
          grepl("cefazolin", DrugName, ignore.case = TRUE) ~ 3,
          grepl("cephalexin", DrugName, ignore.case = TRUE) ~ 3,
          grepl("erythromycin", DrugName, ignore.case = TRUE) ~ 3,
          grepl("penicillin", DrugName, ignore.case = TRUE) ~ 3,
          grepl("dicloxacillin", DrugName, ignore.case = TRUE) ~ 2,
          grepl("nafcillin", DrugName, ignore.case = TRUE) ~ 2,
          grepl("oxacillin", DrugName, ignore.case = TRUE) ~ 2,
          grepl("metronidazole", DrugName, ignore.case = TRUE) ~ 2,
          grepl("tinidazole", DrugName, ignore.case = TRUE) ~ 2,
          TRUE ~ 0)) %>% # NOTE: Optional, but can replace SpectrumScore with institution-specific scoring pattern
      group_by(ExamplePatientID, DrugName, AdminDate, OrderingPhysicianName) %>%
      mutate(DOT = ifelse(row_number() == 1, 1, 0)) %>%
      ungroup() %>%
      group_by(ExamplePatientID, OrderingPhysicianName) %>%
      mutate(UniqueMDEncounter = ifelse(row_number() == 1, 1, 0)) %>%
      ungroup()
    
    return(data)
}

# Filters data for systemic ABX and ID Providers
filter_id_systemic_abx <- function(data) {
  data %>% filter(SystemicABX == 1, IDProvider == 1, DOT == 1) 
}

# Creates pivottable for DOT/encounter and Spectrum Score/Encounter variables
create_spectrum_score_pivot <- function(data) {
  filtered_data <- filter_id_systemic_abx(data)
  pivot_data <- filtered_data %>%
    group_by(OrderingPhysicianName) %>%
    summarize(
      Average_CMI = mean(CMI, na.rm = TRUE),
      Sum_SpectrumScore = sum(SpectrumScore, na.rm = TRUE),
      Sum_DOT = sum(DOT, na.rm = TRUE),
      Sum_UniqueMDEncounter = sum(UniqueMDEncounter, na.rm = TRUE),
      DOT_per_Encounter = sum(DOT, na.rm = TRUE) / sum(UniqueMDEncounter, na.rm = TRUE),
      SpectrumScore_per_Encounter = sum(SpectrumScore, na.rm = TRUE) / sum(UniqueMDEncounter, na.rm = TRUE)) %>%
    arrange(desc(SpectrumScore_per_Encounter))
  return(pivot_data)
}

# Creates pivottable for antibiotics that comprise the DASC for each provider 
create_antibiotic_composition_DASC_pivot <- function(data) {
  filtered_data <- filter_id_systemic_abx(data)
  pivot_data <- filtered_data %>%
    group_by(OrderingPhysicianName, DrugName) %>%
    summarize(SpectrumScore_Sum = sum(SpectrumScore, na.rm = TRUE), .groups = "drop") %>%
    pivot_wider(names_from = DrugName, values_from = SpectrumScore_Sum, values_fill = 0)
  drug_columns <- setdiff(colnames(pivot_data), "OrderingPhysicianName")
  pivot_data <- pivot_data %>%
    mutate(Total_DASC = rowSums(select(., all_of(drug_columns)), na.rm = TRUE)) %>%
    mutate(across(all_of(drug_columns), ~ (. / Total_DASC), .names = " {.col}")) %>%
    select(OrderingPhysicianName, starts_with(" "), Total_DASC) %>%
    arrange(desc(Total_DASC))
  return(pivot_data)
}

# Creates pivottable for antibiotics that comprise the DOT for each provider
create_antibiotic_composition_DOT_pivot <- function(data) {
  filtered_data <- filter_id_systemic_abx(data)
  pivot_data <- filtered_data %>%
    group_by(OrderingPhysicianName, DrugName) %>%
    summarize(DOT_Sum = sum(DOT, na.rm = TRUE), .groups = "drop") %>%
    pivot_wider(
      names_from = DrugName,
      values_from = DOT_Sum,
      values_fill = 0)
  drug_columns <- setdiff(colnames(pivot_data), "OrderingPhysicianName")
  pivot_data <- pivot_data %>%
    mutate(Total_DOT = rowSums(select(., all_of(drug_columns)), na.rm = TRUE)) %>%
    mutate(across(all_of(drug_columns), ~ (. / Total_DOT), .names = " {.col}")) %>%
    select(OrderingPhysicianName, starts_with(" "), Total_DOT) %>%
    arrange(desc(Total_DOT))
  return(pivot_data)
}

# Function to run process to generate columns and create pivottables
run_analysis <- function() {
    data <- load_and_transform_data()
    list(
        spectrum_pivot = create_spectrum_score_pivot(data),
        DASC_pivot = create_antibiotic_composition_DASC_pivot(data),
        DOT_pivot = create_antibiotic_composition_DOT_pivot(data))
}

# Run analysis and prompt user to select Excel file with the adata
cat("Please click on the Excel file with antibiotic administration data...\n")
pivot_tables <- run_analysis()

# Defining variables from pivottables to read data 
spectrum_pivot <- pivot_tables$spectrum_pivot
DASC_pivot <- pivot_tables$DASC_pivot
DOT_pivot <- pivot_tables$DOT_pivot

data_1 <- spectrum_pivot   
data_2 <- DASC_pivot       
data_3 <- DOT_pivot        

# Extracts and de-identifies IDProviders with random letters
generate_identifiers <- function(n) {
    identifiers <- c(LETTERS, outer(LETTERS, LETTERS, paste0))  
    return(identifiers[1:n])
}

de_identify_providers <- function(data_1, data_2, data_3) {
    all_providers <- unique(c(
        na.omit(data_1$OrderingPhysicianName),
        na.omit(data_2$OrderingPhysicianName),
        na.omit(data_3$OrderingPhysicianName)))
    
    identifiers <- generate_identifiers(length(all_providers))
    provider_mapping <- data.frame(
        OriginalName = all_providers,
        Identifier = identifiers,
        stringsAsFactors = FALSE)

    replace_providers <- function(df, column_name) {
        df[[column_name]] <- provider_mapping$Identifier[match(df[[column_name]], provider_mapping$OriginalName)]
        return(df)
    }
    
    data_1 <- replace_providers(data_1, "OrderingPhysicianName")
    data_2 <- replace_providers(data_2, "OrderingPhysicianName")
    data_3 <- replace_providers(data_3, "OrderingPhysicianName")
    
    # Return de-identified datasets and mapping
    return(list(
        data_1 = data_1,
        data_2 = data_2,
        data_3 = data_3,
        mapping = provider_mapping
    ))
}

deidentified_data <- de_identify_providers(data_1, data_2, data_3)

# Replace the original datasets with the de-identified versions
data_1 <- deidentified_data$data_1
data_2 <- deidentified_data$data_2
data_3 <- deidentified_data$data_3

# Prompting directory to both save report cards into as well as the mapping legend
cat("Please choose a folder to save report cards into...\n")
save_dir <- choose.dir(caption = "Select a folder to save report cards") #NOTE: Optional, but this can be modified to a hard-coded file path if easier
mapping_filename <- file.path(save_dir, "Provider_Identifier_Mapping_Legend.xlsx")
writexl::write_xlsx(deidentified_data$mapping, mapping_filename)

# Rename columns and extract relevant columns
colnames(data_1) <- c("Provider", "CMI_average", "SpectrumScore", "DOT", 
                                            "UniqueMDEncounter", "DOT_per_Encounter", "SpectrumScore_per_Encounter")
data_subset <- data_1[, c("Provider", "CMI_average", "SpectrumScore_per_Encounter", 
                                                    "DOT_per_Encounter", "UniqueMDEncounter")]

# Converting both DASC and DOT composition to percentages 
data_long <- data_2 %>%
    pivot_longer(cols = -`OrderingPhysicianName`, names_to = "Antibiotic", values_to = "Percentage") %>%
    rename(Provider = `OrderingPhysicianName`) %>%
    mutate(Percentage = as.numeric(Percentage))  
data_3_long <- data_3 %>%
    pivot_longer(cols = -`OrderingPhysicianName`, names_to = "Antibiotic", values_to = "Percentage") %>%
    rename(Provider = `OrderingPhysicianName`) %>%
    mutate(Percentage = as.numeric(Percentage))  
data_long <- data_long %>%
    mutate(Percentage = round(Percentage * 100, 1))  
data_3_long <- data_3_long %>%
    mutate(Percentage = round(Percentage * 100, 1))  

# Identify top 5 antibiotics for both DASC and DOT (skipping the total_DASC and total_DOT)
top_5_antibiotics <- data_long %>%
    group_by(Provider) %>%
    arrange(Provider, desc(Percentage)) %>%
    slice(2:6) %>%
    ungroup()

top_5_dot <- data_3_long %>%
    group_by(Provider) %>%
    arrange(Provider, desc(Percentage)) %>%
    slice(2:6) %>%
    ungroup()

    # Remove rows with missing values (if any) and excluding rows with CMI = 0 or outside 95th percentile for color/range calculation
    data_subset <- na.omit(data_subset)
    data_subset_no_zero_cmi <- data_subset[data_subset$CMI_average != 0, ]
    cmi_cap <- quantile(data_subset_no_zero_cmi$CMI_average, 0.95, na.rm = TRUE)
    data_subset$CMI_average_capped <- pmin(data_subset$CMI_average, cmi_cap)
    data_subset$CMI_color <- ifelse(data_subset$CMI_average == 0, NA, data_subset$CMI_average_capped)

    # Compute the Minimum Covariance Determinant and inliers/center for ellipse
    mcd_result <- cov.mcd(data_subset[, c("SpectrumScore_per_Encounter", "DOT_per_Encounter")])
    data_subset$is_inlier <- FALSE  
    data_subset$is_inlier[mcd_result$best] <- TRUE  
    inliers <- data_subset[data_subset$is_inlier == TRUE, ]
    ellipse_center <- colMeans(data_subset[mcd_result$best, c("SpectrumScore_per_Encounter", "DOT_per_Encounter")])

    # Generate personalized report cards for each provider - starting with higlighted graph
    for (provider in unique(data_subset$Provider)) {
      data_subset$highlight <- ifelse(data_subset$Provider == provider, "Highlighted", "Other")

      # Adding Provider name and coordinates
      data_subset$ProviderLabel <- ifelse(data_subset$Provider == provider, provider, "")
      highlighted_coords <- data_subset[data_subset$Provider == provider, c("SpectrumScore_per_Encounter", "DOT_per_Encounter")]

     

      # Applying color gradient to CMI values (if NA or 0, will not use color gradient)
      if (all(is.na(data_subset$CMI_average_capped) | data_subset$CMI_average_capped == 0)) {
        color_layer_highlight <- geom_point(aes(size = UniqueMDEncounter, 
                                                alpha = ifelse(highlight == "Highlighted", 0.6, 0.1)))
      } else {
        color_layer_highlight <- list(
          geom_point(aes(size = UniqueMDEncounter, 
                         color = ifelse(CMI_average_capped == 0, NA, CMI_average_capped),
                         alpha = ifelse(highlight == "Highlighted", 0.6, 0.1))),
          scale_color_gradient(low = "green", high = "red", name = "Average Case Mix Index (CMI)")
        )
      }
      # Creating the plot, x-axis and y-axis in plots dyanmically determined based on sample 
      plot_highlighted <- ggplot(data_subset, aes(x = SpectrumScore_per_Encounter, y = DOT_per_Encounter))
      x_max <- max(data_subset$SpectrumScore_per_Encounter, na.rm = TRUE)
      x_max <- ceiling(x_max * 1.05) 
      y_max <- max(data_subset$DOT_per_Encounter, na.rm = TRUE)
      y_max <- ceiling(y_max * 1.05) 

      # Generating highlighted plot
      plot_highlighted <- plot_highlighted +
        color_layer_highlight +
        scale_size_continuous(name = "Unique Patient Encounters", range = c(8, 25)) +
        scale_alpha_continuous(range = c(0.2, 1), guide = "none") +
        stat_ellipse(data = inliers, aes(x = SpectrumScore_per_Encounter, y = DOT_per_Encounter), 
                     level = 0.95, color = "black", size = 1) +
        annotate("point", x = ellipse_center[1], y = ellipse_center[2], shape = 3, color = "blue", size = 7, stroke = 1.5) +
        annotate("text", x = ellipse_center[1], y = ellipse_center[2] - 4, # NOTE: these coordinates for the center may need to be manually adjusted based on the sample
                 label = paste0("Center: (", round(ellipse_center[1], 0), ", ", 
                                round(ellipse_center[2], 0), ")"),
                 size = 5, color = "blue", hjust = 0.5) +
        annotate("segment", x = ellipse_center[1], y = ellipse_center[2], 
                 xend = ellipse_center[1], yend = ellipse_center[2] - 3, 
                 color = "blue", linetype = "dashed", size = 0.8) +  
        annotate("text", x = highlighted_coords$SpectrumScore_per_Encounter, 
                 y = highlighted_coords$DOT_per_Encounter + 1,  
                 label = paste0("(", round(highlighted_coords$SpectrumScore_per_Encounter, 0), ", ",
                                round(highlighted_coords$DOT_per_Encounter, 0), ")"), 
                 size = 5, color = "blue") +
        geom_text(aes(label = ProviderLabel), 
                  size = 5, vjust = 0.5,  hjust = 0.5,   show.legend = FALSE) +  
        scale_y_continuous(breaks = seq(0, y_max, by = 1), limits = c(0, y_max), expand = expansion(mult = c(0, 0.02))) +
        scale_x_continuous(limits = c(0, x_max), expand = expansion(mult = c(0, 0.02))) +
        ggtitle(paste("Antibiotic Utilization Data from January 2024 through December 2024 \nSpecific Prescriber Report for", provider)) +
        xlab("Days of Antibiotic Spectrum Coverage (DASC) per Encounter") +
        ylab("Days of Therapy (DOT) per Encounter") +
        theme_minimal() +

      # Additional customization of graph
        theme(
          legend.position = "right",     
          panel.grid.major.y = element_line(size = 0.5, color = "grey"),  
          panel.grid.minor.y = element_blank(),  
          axis.text.x = element_text(size = 14),  
          axis.text.y = element_text(size = 14),  
          axis.title.x = element_text(size = 16, face = "bold"),  
          axis.title.y = element_text(size = 16, face = "bold"),  
          plot.title = element_text(size = 18, face = "bold"),  
          axis.ticks.y = element_line(color = "black"),  
          axis.ticks.x = element_line(color = "black"),   
          plot.caption = element_text(size = 16, hjust = 1, vjust = 1)  
        ) +
        guides(color = guide_colorbar(title = "CMI Average"),
               size = guide_legend(title = "Unique MD Encounters")) +
        labs(caption = "Ellipse represents 95% percentile of covariance matrix")

      # Save the highlighted plot to the prior select directory as a png
      plot_file_highlighted <- file.path(save_dir, paste0(gsub("\\*", "", provider), "_highlighted_report_card.png"))
      ggsave(filename = plot_file_highlighted, plot = plot_highlighted, width = 12, height = 8, dpi = 300)

      # Create another plot without highlighting the provider
      plot_no_highlight <- ggplot(data_subset, aes(x = SpectrumScore_per_Encounter, y = DOT_per_Encounter))

      # Similar to before, applying color gradient to CMI values (if NA or 0, will not use color gradient)
      if (all(is.na(data_subset$CMI_average_capped) | data_subset$CMI_average_capped == 0)) {
        color_layer <- geom_point(aes(size = UniqueMDEncounter, 
                                      color = ifelse(CMI_average_capped == 0, NA, data_subset$CMI_average_capped),
                                      alpha = 0.2))
      } else {
        color_layer <- list(
          geom_point(aes(size = UniqueMDEncounter, 
                         color = ifelse(CMI_average_capped == 0, NA, CMI_average_capped),
                         alpha = 0.2)),
          scale_color_gradient(low = "green", high = "red", name = "Average Case Mix Index (CMI)")
        )
      }

      # Generating non-highlighted plot
      plot_no_highlight <- plot_no_highlight + 
        color_layer + 
        scale_size_continuous(name = "Unique Patient Encounters", range = c(8, 25)) +
        scale_alpha_continuous(range = c(0.2, 1), guide = "none") +
        stat_ellipse(data = inliers, aes(x = SpectrumScore_per_Encounter, y = DOT_per_Encounter), 
                     level = 0.95, color = "black", size = 1) +
        annotate("point", x = ellipse_center[1], y = ellipse_center[2], shape = 3, color = "blue", size = 7, stroke = 1.5) +
        annotate("text", x = ellipse_center[1], y = ellipse_center[2] - 4, 
                 label = paste0("Center: (", round(ellipse_center[1], 0), ", ", 
                                round(ellipse_center[2], 0), ")"),
                 size = 5, color = "blue", hjust = 0.5) +
        annotate("segment", x = ellipse_center[1], y = ellipse_center[2], 
                 xend = ellipse_center[1], yend = ellipse_center[2] - 3, 
                 color = "blue", linetype = "dashed", size = 0.8) +
        geom_text(aes(label = Provider), 
                  size = 5, vjust = 0.5,  hjust = 0.5,   show.legend = FALSE) +  
        scale_y_continuous(breaks = seq(0, y_max, by = 1), limits = c(0, y_max), expand = expansion(mult = c(0, 0.02))) +
        scale_x_continuous(limits = c(0, x_max), expand = expansion(mult = c(0, 0.02))) +
        ggtitle("Antibiotic Utilization Data\n ID Prescribers") +
        xlab("Days of Antibiotic Spectrum Coverage (DASC) per Encounter") +
        ylab("Days of Therapy (DOT) per Encounter") +
        theme_minimal() +

        # Additional customization of plot
        theme(
          legend.position = "right",     
          panel.grid.major.y = element_line(size = 0.5, color = "grey"),  
          panel.grid.minor.y = element_blank(),  
          axis.text.x = element_text(size = 14),  
          axis.text.y = element_text(size = 14),  
          axis.title.x = element_text(size = 16, face = "bold"),  
          axis.title.y = element_text(size = 16, face = "bold"),  
          plot.title = element_text(size = 18, face = "bold"),  
          axis.ticks.y = element_line(color = "black"),  
          axis.ticks.x = element_line(color = "black"), 
          plot.caption = element_text(size = 16, hjust = 1, vjust = 1) 
        ) +
        guides(color = guide_colorbar(title = "CMI Average"),
               size = guide_legend(title = "Unique MD Encounters")) +
        labs(caption = "Ellipse represents 95% percentile of covariance matrix")

      # Save the non-highlighted plot to prior directory
      plot_file_no_highlight <- file.path(save_dir, "no_highlight_report_card.png")
      ggsave(filename = plot_file_no_highlight, plot = plot_no_highlight, width = 12, height = 8, dpi = 300)
    

# Create a Word document to start the autogenerated Word document process 
# NOTE: all text and customization can be modified as per institution preference
doc <- read_docx() %>%
body_add_fpar(
  fpar(
    ftext(
      paste0("Antibiotic Stewardship Report - Prescriber ", provider),
      prop = fp_text(font.size = 20, font.family = "Franklin Gothic Book", bold = TRUE, underline = TRUE)
    )
  )
)%>%
  body_add_fpar(
    fpar(ftext("Antimicrobial Stewardship Team", 
               prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book Book", italic = TRUE)))
  ) %>%
  body_add_fpar(
    fpar(ftext("ID Prescriber Data at our Institution", 
               prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book Book", italic = TRUE)))
  ) %>%
  body_add_fpar(fpar(ftext("", prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book"))))


# Add background section
doc <- doc %>%
  body_add_fpar(
    fpar(ftext("Background", prop = fp_text(bold = TRUE, font.size = 16, font.family = "Franklin Gothic Book")))
  ) %>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "The 2019 CDC Core Elements of Hospital Antibiotic Stewardship recognizes prospective audit and feedback as a core element of an antibiotic stewardship program. Similarly, the 2007 IDSA/SHEA guidelines emphasize that prospective audit and feedback form the foundation of a stewardship program.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") 

doc <- doc %>%
  body_add_fpar(
    fpar(ftext("What is included in this report?", 
               prop = fp_text(bold = TRUE, font.size = 16, font.family = "Franklin Gothic Book")))
  ) %>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "The Antimicrobial Stewardship Team has reviewed your individual antibiotic prescribing patterns, including:",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
  body_add_fpar(fpar(ftext("   • duration of therapy (measured via Days of Therapy)", 
                       prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")))) %>%
  body_add_fpar(fpar(ftext("   • spectrum of activity (measured via Days of Antibiotic Spectrum Therapy as described on next page)", 
                       prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")))) %>%
  body_add_fpar(fpar(ftext("   • de-identified comparison to other ID prescribers", 
                       prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")))) %>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "Patient volume (number of unique encounters) was also incorporated to present a comprehensive picture. Prescribers within the ellipse represent inliers of typical prescribing patterns among your peers.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "Please, take a moment to review the information in this document for assess your antibiotic utilization performance and prescribing habits.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  )

doc <- body_add_break(doc)

# Add spectrum score background
doc <- doc %>%
  body_add_fpar(
    fpar(ftext("Using Spectrum Score to help assess antibiotic utilization",
               prop = fp_text(bold = TRUE, font.size = 16, font.family = "Franklin Gothic Book")))
  ) %>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "Spectrum scores are a relatively novel metric for assessing the appropriateness of antibiotic utilization and address some limitations of traditional metrics. These scoring systems aim to objectively quantify antibiotic activity based on the spectrum of activity against common pathogens. Higher scores represent broader antibiotic utilization.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  )%>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "In this report, we have further optimized the Days of Antibiotic Spectrum Coverage (DASC) score described by Kakiuchi et al. (Clinical Infectious Diseases, 2022), 
to provide a more meaningful assessment in the context of inpatient antibiotic stewardship by taking antibiotic resistant organisms and mechanisms of resistance into account.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  )%>%
body_add_par("", style = "Normal") %>%
  body_add_fpar(
    fpar(ftext(
      "Please see below graph for a peer comparison visual of the prescribing habits of all the ID prescribers within our institution. 
The x-axis represents the spectrum of activity prescribed and the y-axis represents the duration prescribed.",
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  )%>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") %>%
body_add_par("", style = "Normal") 

# Add the non-highlighted plot to the document
  doc <- doc %>% 
    body_add_img(src = plot_file_no_highlight, width = 6, height = 4) %>% 
    body_add_par("", style = "Normal")


doc <- body_add_break(doc)

# Add the highlighted plot to the document
  doc <- doc %>% 
    body_add_img(src = plot_file_highlighted, width = 6, height = 4) %>% 
    body_add_par("", style = "Normal")

# Add summary bullet points of provider performance
doc <- doc %>%
  body_add_fpar(
    fpar(ftext(
      paste0("You are Prescriber ", provider, "."),
      prop = fp_text(font.size = 16, font.family = "Franklin Gothic Book", bold = TRUE, underline = TRUE, color = "darkred")
    ))
  ) %>%
  body_add_fpar(
    fpar(ftext(
      if (data_subset$is_inlier[data_subset$Provider == provider]) {
        "• Your antibiotic utilization is similar to those of your peers."
      } else {
        "• Your antibiotic utilization is an outlier compared to your peers, which may indicate opportunities to refine your antibiotic selection. Table 1 and 2 highlight your most commonly prescribed antibiotics."
      },
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
  body_add_fpar(
    fpar(ftext(
      paste0(
        "• Compared to your peers, ",
        round(percent_rank(data_subset$SpectrumScore_per_Encounter)[data_subset$Provider == provider] * 100, 1),
        "% prescribed narrower-spectrum antibiotics than you."
      ),
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
  body_add_fpar(
    fpar(ftext(
      paste0(
        "• Compared to your peers, ",
        round(percent_rank(data_subset$DOT_per_Encounter)[data_subset$Provider == provider] * 100, 1),
        "% prescribed shorter durations of antibiotics than you."
      ),
      prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book")
    ))
  ) %>%
  body_add_par("", style = "Normal")  

# Prepare antibiotic composition for specific providers
provider_data_antibiotics <- top_5_antibiotics %>% 
  filter(Provider == provider) %>%
  mutate(Percentage = paste0(Percentage, "%"))%>%
  rename(Prescriber = Provider)

provider_data_dot <- top_5_dot %>% 
  filter(Provider == provider) %>%
  mutate(Percentage = paste0(Percentage, "%"))%>%
  rename(Prescriber = Provider)

# Add the antibiotics tables for DASC and DOT for the specific provider
doc <- doc %>%
  body_add_fpar(fpar(ftext("Top 5 Antibiotics Contributing to your Antibiotic Spectrum Coverage Score:", prop = fp_text(bold = TRUE, font.size = 14, font.family = "Franklin Gothic Book")))) %>%
  body_add_table(provider_data_antibiotics, style = "table_template") %>%
  body_add_par("", style = "Normal")

doc <- doc %>%
  body_add_fpar(fpar(ftext("Top 5 Antibiotics Contributing to your Days of Therapy (DOT):", prop = fp_text(bold = TRUE, font.size = 14, font.family = "Franklin Gothic Book")))) %>%
  body_add_table(provider_data_dot, style = "table_template") %>%
  body_add_par("", style = "Normal")

# Summary paragraph
doc <- doc %>%
  body_add_fpar(fpar(ftext("If you have any questions, feedback, or comments for this report, please reach out to any member of the Antimicrobial Stewardship Team.", prop = fp_text(font.size = 12, font.family = "Franklin Gothic Book", italic = TRUE)))) %>%
  body_add_par("", style = "Normal")

# Save the Word document to prior directory
report_file <- file.path(save_dir, paste0(gsub("\\*", "", provider), "_Stewardship_Report_2024.docx"))
print(doc, target = report_file)
}
