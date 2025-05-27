install.packages("forecast")###
library(forecast)
library(moments) #for skewness
library(tidyverse) ####
library(scales)###
library(dplyr)
install.packages("writexl")
library(writexl)
library(readxl)
sales <- read_excel("Business Analytics/MISM6202/Analytics Problem Set/sales.xlsx", 
                    sheet = "sales-data")
View(sales)
# Create the summary table
customer_level_transaction <- sales%>%
  filter(!is.na(customer.id)) %>%  # Exclude rows with missing customer IDs
  group_by(customer.id) %>%
  summarise(
    total_items_purchased = sum(qty, na.rm = TRUE), # Total number of items purchased
    avg_item_sale_price = mean(sale.amount / qty, na.rm = TRUE) # Average item sale price
  )
print(customer_level_transaction, n = Inf)  
View(customer_level_transaction)

customers_cleaned<- read_excel("Business Analytics/MISM6202/Analytics Problem Set/customers_cleaned_KC.xlsx", 
                                   sheet = "customers")
View(customers_cleaned)

customer_purchases <- customer_level_transaction %>%
  inner_join(customers_cleaned, by = "customer.id") 
write.csv(customer_purchases, "customer_purchases.csv", row.names = FALSE)
write_xlsx(customer_purchases, "customer_purchases.xlsx")
path <- getwd()
text <- "The files are saved at"
print(paste(text, path))

mean_sale_amount <- mean(sales$sale.amount)
median_sale_amount <- median(sales$sale.amount)
std_dev_sale_amount <- sd(sales$sale.amount)
skewness_sale_amount <- skewness(sales$sale.amount)

print(paste("Sale Amount Statistics:"))
print(paste("Mean:", mean_sale_amount))
print(paste("Median:",median_sale_amount))
print(paste("Standard Deviation:", std_dev_sale_amount))
print(paste("Skewness Coefficient:", skewness_sale_amount))

summary(sales$sale.amount)

summary_stats_table <- data.frame(
  Statistic = c("Mean", "Median", "Standard Deviation", "Skewness Coefficient"),
  Value = c(mean_sale_amount, median_sale_amount, std_dev_sale_amount, 
            skewness_sale_amount)
)
print(summary_stats_table)


#boxplot All Sales
boxplot_all_sales <- boxplot(sales$sale.amount, data=sales, main = "Boxplot of Sale Amount for All Sales Amount",
                             ylab= "Sale Amount", horizontal = TRUE)

#boxplot for each category
par(mar = c(4, 10, 2, 2))
boxplot_prod_cat <- boxplot(sales$sale.amount~sales$category, data=sales, 
        main="Boxplot of Sale Amount for Each Category",
        xlab= "Sale Amount", ylab="", 
        col = rainbow(length(unique(sales$category))),
        las = 1,
        outline = TRUE,
        horizontal=TRUE)

par(mar = c(5, 4, 4, 2) + 0.1)

#Blended Gross Margin for each category
df_accessories <- subset(sales, category == "Accessories")
blended_gm_accessories <- (sum(df_accessories$sale.amount,na.rm=FALSE)-sum(df_accessories$ext.cost))/sum(df_accessories$sale.amount)
blended_gm_accessories

df_childrens <- subset(sales, category == "Childrens")
blended_gm_childrens <- (sum(df_childrens$sale.amount,na.rm=FALSE)-sum(df_childrens$ext.cost))/sum(df_childrens$sale.amount)
blended_gm_childrens

df_gifts <- subset(sales, category == "Gifts & Lifestyle")
blended_gm_gifts <- (sum(df_gifts$sale.amount,na.rm=FALSE)-sum(df_gifts$ext.cost))/sum(df_gifts$sale.amount)
blended_gm_gifts

df_footwear <- subset(sales, category == "Footwear")
blended_gm_footwear <- (sum(df_footwear$sale.amount,na.rm=FALSE)-sum(df_footwear$ext.cost))/sum(df_footwear$sale.amount)
blended_gm_footwear

df_mens <- subset(sales, category == "Men's Apparel")
blended_gm_mens <- (sum(df_mens$sale.amount,na.rm=FALSE)-sum(df_mens$ext.cost))/sum(df_mens$sale.amount)
blended_gm_mens

df_womens <- subset(sales, category == "Women's Apparel")
blended_gm_womens <- (sum(df_womens$sale.amount,na.rm=FALSE)-sum(df_womens$ext.cost))/sum(df_womens$sale.amount)
blended_gm_womens

#Outliers using boxplot method
outliers_all <- boxplot(sales$sale.amount)$out
outliers_all
length(outliers_all)

#Outliers using z-score method
# Calculate the z-scores for sale.amount
sales$z_score <- (sales$sale.amount - mean(sales$sale.amount, na.rm = TRUE)) /
  sd(sales$sale.amount, na.rm = TRUE)

# Identify outliers with z-score > 3 or < -3
outliers_z_score <- sales %>%
  filter(abs(z_score) > 3)

# View outliers
outliers_z_score$sale.amount
length(outliers_z_score$sale.amount)


#Hypothesis Testing
# Perform the ANOVA test
anova_result <- aov(gross.margin ~ price.category, data = sales)

# Summarize the results of the ANOVA
summary(anova_result)

#Regression
# Fit a linear regression model
model <- lm(gross.margin ~ sale.amount + price.category + qty + unit.cost + store + category + loyalty.member, data = sales)

# Summary of the model
summary(model)


