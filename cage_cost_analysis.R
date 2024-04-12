rm(list = ls())

required_packages <- c("readxl", "dplyr", "lubridate", "ggplot2")
for (package in required_packages) {
  if (!(package %in% installed.packages())) {
    install.packages(package)
  }
}

library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)

df <- read_excel("CagesFrom2023.xlsx")

df$`Creation D.` <- as.Date(df$`Creation D.`)
df$`Elimination D.` <- as.Date(df$`Elimination D.`)

cage_costs <- c("GM500B" = 9.75, "GM500" = 7.25, "GM500S" = 9.75)

daily_costs <- data.frame(Date = as.Date(character()), Cage_Type = character(), Cost = numeric())

for (i in 1:nrow(df)) {
  row <- df[i, ]
  creation_date <- row$`Creation D.`
  elimination_date <- row$`Elimination D.`
  cage_type <- strsplit(as.character(row$`Cage type`), " ")[[1]][1]
  
  if (!(cage_type %in% names(cage_costs))) {
    next
  }
  
  num_days <- as.numeric(elimination_date - creation_date) + 1
  
  dates_used <- seq(creation_date, by = "day", length.out = num_days)
  
  for (date in dates_used) {
    daily_costs <- rbind(daily_costs, data.frame(Date = date, Cage_Type = cage_type, Cost = cage_costs[cage_type]))
  }
}

daily_costs_summary <- daily_costs %>%
  group_by(Date, Cage_Type) %>%
  summarise(Total_Cost = sum(Cost))

daily_costs_summary$Date <- as.Date(daily_costs_summary$Date)

ggplot(daily_costs_summary, aes(x = Date, y = Total_Cost, color = Cage_Type)) +
  geom_line() +
  labs(x = "Date", y = "Total Cost", title = "Daily Costs of Different Cage Types") +
  scale_x_date(date_labels = "%Y-%m-%d", date_breaks = "1 month") +  
  theme_minimal()

