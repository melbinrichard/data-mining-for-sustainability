# Load necessary libraries
library(MASS)
library(dplyr)
library(readxl)
library(caret)

# Read the Excel file
data <- read_excel("BP_extracted_financial_data_Final.xlsx")

# View the structure of the data
str(data)

# Standardize the data
preproc <- preProcess(data, method = c("center", "scale"))
standardized_data <- predict(preproc, data)

# View standardized data
head(standardized_data)

# Fit a multivariate linear model - Aim1
mlm_model <- lm(cbind(Fin1, Fin2, Fin3, Fin4, Fin5, Fin6, Fin7, Fin8, Fin9, Fin10) ~ Aim1, data = standardized_data)

# Fit a multivariate linear model - Aim2
mlm_model <- lm(cbind(Fin1, Fin2, Fin3, Fin4, Fin5, Fin6, Fin7, Fin8, Fin9, Fin10) ~ Aim2, data = standardized_data)

# Fit a multivariate linear model - Aim3
mlm_model <- lm(cbind(Fin1, Fin2, Fin3, Fin4, Fin5, Fin6, Fin7, Fin8, Fin9, Fin10) ~ Aim3, data = standardized_data)

# Fit a multivariate linear model - Aim4
mlm_model <- lm(cbind(Fin1, Fin2, Fin3, Fin4, Fin5, Fin6, Fin7, Fin8, Fin9, Fin10) ~ Aim4, data = standardized_data)

# Fit a multivariate linear model - Aim5
mlm_model <- lm(cbind(Fin1, Fin2, Fin3, Fin4, Fin5, Fin6, Fin7, Fin8, Fin9, Fin10) ~ Aim5, data = standardized_data)

# View summary of the multivariate regression model
summary(mlm_model)
