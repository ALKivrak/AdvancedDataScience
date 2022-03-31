install.packages("ggplot2")
install.packages("dplyr")
install.packages("naivebayes")
install.packages("InformationValue")
library(naivebayes)
library(dplyr)
library(ggplot2)
library(InformationValue)

data <- read.csv(file = "https://github.com/Violagameboy/AdvancedDataScience/blob/gh-pages/diabetes.csv", header = T)
str(data)

#logistic Regression

logistic <- glm(Outcome ~ Age, data = train, family = "binomial")
options(scipen=999)
summary(logistic)
ldata <- data.frame(Age=seq(min(data$Age), max(data$Age),len=5))
ldata$Outcome = predict(logistic, ldata, type="response")

plot(Outcome ~ Age, data=data, col="steelblue")
lines(Outcome ~ Age, ldata, lwd=2)

ggplot(data, aes(x=Age,y=Outcome)) +
  geom_point(alpha=.5) +
  stat_smooth(method="glm", se=FALSE, method.args = list(family=binomial))
#naive bayes
data$Outcome <- as.factor(data$Outcome)
set.seed(1234)
ind <- sample(2, nrow(data), replace = T, prob = c(0.8, 0.2))
train <- data[ind == 1,]
test <- data[ind == 2,]
ndata <- naive_bayes(Outcome ~ ., data = train, usekernel = T)
p <- predict(ndata, test)
plot(ndata)
table(p, test$Outcome)
