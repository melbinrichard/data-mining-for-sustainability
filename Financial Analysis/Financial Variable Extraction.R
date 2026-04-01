# Loading the libraries
library(readxl)
library(dplyr)

# Loading the finance data
BP_Finance <- "BP_2023_Financial_Data.xlsx"

# Group Income Statement sheet - Extract Revenue, Operating Expenses
income_statement <- read_excel(BP_Finance, sheet = "Group income statement")
income_statement <- income_statement %>%  filter(Type %in% c("Total revenues and other income", "Production and manufacturing expenses")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Group cash flow statement sheet - Extract Cash Flow from Operations, Capital Expenditure
cash_flow <- read_excel(BP_Finance, sheet = "Group cash flow statement")
cash_flow <- cash_flow %>%  filter(Type %in% c("Net cash provided by operating activities", "Total cash capital expenditure")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Group balance sheet - Extract Total Assets, Total Liabilities
balance_sheet <- read_excel(BP_Finance, sheet = "Group balance sheet")
balance_sheet <- balance_sheet %>%  filter(Type %in% c("Net assets", "Total liabilities")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Net debt & net debt incl leases Sheet - Extract Net Debt
net_debt <- read_excel(BP_Finance, sheet = "Net debt & net debt incl leases")
net_debt <- net_debt %>%  filter(Type %in% c("Net debt including leases")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Capital exp cash basis Sheer - Extract Capital Expenditure in Renewable Energy
capex_renewable <- read_excel(BP_Finance, sheet = "Capital exp cash basis")
capex_renewable <- capex_renewable %>%  filter(Type %in% c("Total")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Gas & Low Carbon Energy Sheet - Extract Investment in Low-Carbon Projects
low_carbon <- read_excel(BP_Finance, sheet = "Gas & Low Carbon Energy")
low_carbon <- low_carbon %>%  filter(Type %in% c("Investment in Low-Carbon Projects")) %>% select(Type, '2019', '2020', '2021', '2022', '2023')

# Summary Sheet - Extract ROACE%
summary <- read_excel(BP_Finance, sheet = "Summary")
summary <- summary %>%  filter(Type %in% c("ROACE%")) %>% select(Type, '2019', '2020', '2021', '2022', '2023') %>% mutate_at(vars('2019', '2020', '2021', '2022', '2023'), as.numeric)

# Combining to a Single Table 
final_table <- bind_rows(income_statement, cash_flow, balance_sheet, net_debt, capex_renewable, low_carbon, summary)
final_table <- final_table %>%  mutate(Type = recode(Type, "Total revenues and other income" = "Total Revenue", "Production and manufacturing expenses" = "Operating Expenses", "Net cash provided by operating activities" = "Cash Flow from Operations", "Total cash capital expenditure" = "Capital Expenditure", "Net assets" = "Total Assets", "Net debt including leases" = "Net Debt", "Total" = "Capital Expenditure in Renewable Energy" )) %>% mutate_at(vars('2019', '2020', '2021', '2022', '2023'), as.numeric)
colnames(final_table)[colnames(final_table) == "Type"] <- "Key Financial Indicators"
print(final_table)

# Saving the data as CSV file
write.csv(final_table, "extracted_financial_data.csv", row.names = FALSE)
