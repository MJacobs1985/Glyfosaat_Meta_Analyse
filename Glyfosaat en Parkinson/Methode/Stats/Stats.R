#### CLEAR ENVIRONMENT -----
rm(list = ls())
#### CHANGE OPTIONS ----
#options(digits=10)
options(scipen=999)
#### LOAD LIBRARIES -----
library(readxl)
library(tidyverse)
library(Hmisc)
library(tidytext)
library(stringr)
library(metafor)
library(meta)
library(dmetar)

#### LOAD AND WRANGLE DATA ----
df <- read_excel("~/MSJ/Projects/2025/STAF/Glyfosaat en Parkinson/Methode/Data/Data_extracted.xlsx")
colnames(df) = gsub(" ", "_", colnames(df))
colnames(df)
sapply(df, class)

df<-df%>%
  mutate(Risk_Estimate = as.numeric (Risk_Estimate), 
         Risk_Estimate_Low = as.numeric (Risk_Estimate_Low), 
         Risk_Estimate_High = as.numeric (Risk_Estimate_High), 
         Study_Author_Year = paste(Study_Author,Study_Year, sep=","), 
         Total_N = as.numeric(Study_Risk_Patients) + as.numeric(Study_Risk_Controls), 
         Risk_Estimate_Difference = Risk_Estimate_High - Risk_Estimate_Low, 
         Study_Year = as.numeric(Study_Year), 
         Study_Adjustments_Binary = case_when(is.na(Adjustments) ~ 'No', 
                                              TRUE ~ 'Yes'))

#### TABULAR EXPLORATION -----
DataExplorer::introduce(df)
colnames(df)

df%>%
  summarise_all(list(n_distinct))%>%
  unlist()

table(df$Study_Author_Year)
table(df$Risk_Metric)

df%>%
  select(Study_Author_Year, Risk_Metric)

df%>%
  group_by(Study_Design)%>%
  select(Risk_Metric)%>%
  count(Risk_Metric)%>%
  print(n=26)

df%>%
  select(Review_Risk_Metric)%>%
  count(Review_Risk_Metric)%>%
  arrange(desc(n))

df%>%
  select(Risk_Metric, Risk_Estimate)%>%
  drop_na()%>%
  group_by(Risk_Metric)%>%
  summarise(mean = mean(Risk_Estimate, na.rm=TRUE), 
            sd= sd(Risk_Estimate, na.rm=TRUE),  
            var= var(Risk_Estimate, na.rm=TRUE), 
            min = min(Risk_Estimate, na.rm=TRUE),  
            q10 = quantile(Risk_Estimate, probs = 0.1, na.rm=TRUE), 
            q90 = quantile(Risk_Estimate, probs = 0.9, na.rm=TRUE), 
            max= max(Risk_Estimate, na.rm=TRUE))

df%>%
  select(Risk_Metric, Risk_Estimate_Low)%>%
  drop_na()%>%
  group_by(Risk_Metric)%>%
  summarise(mean = mean(Risk_Estimate_Low, na.rm=TRUE), 
            sd= sd(Risk_Estimate_Low, na.rm=TRUE),  
            var= var(Risk_Estimate_Low, na.rm=TRUE), 
            min = min(Risk_Estimate_Low, na.rm=TRUE),  
            q10 = quantile(Risk_Estimate_Low, probs = 0.1, na.rm=TRUE), 
            q90 = quantile(Risk_Estimate_Low, probs = 0.9, na.rm=TRUE), 
            max= max(Risk_Estimate_Low, na.rm=TRUE))

df%>%
  select(Risk_Metric, Risk_Estimate_High)%>%
  drop_na()%>%
  group_by(Risk_Metric)%>%
  summarise(mean = mean(Risk_Estimate_High, na.rm=TRUE), 
            sd= sd(Risk_Estimate_High, na.rm=TRUE),  
            var= var(Risk_Estimate_High, na.rm=TRUE), 
            min = min(Risk_Estimate_High, na.rm=TRUE),  
            q10 = quantile(Risk_Estimate_High, probs = 0.1, na.rm=TRUE), 
            q90 = quantile(Risk_Estimate_High, probs = 0.9, na.rm=TRUE), 
            max= max(Risk_Estimate_High, na.rm=TRUE))

df%>%
  select(Study_Adjustments_Binary)%>%
  table()


df%>%
  mutate(colB = tolower(Adjustments))%>%
  select(colB, Study_Author_Year)%>%
  add_rownames(var = "id")%>%
  drop_na()%>%
  unnest_tokens(Confounder, colB, token = stringr::str_split, pattern = ",")%>%
  mutate(Confounder = stringr::str_squish(Confounder))%>%
  drop_na()%>%
  distinct()%>%
  count(Confounder)%>%
  arrange(desc(n))%>%
  print(n=200)


#### MISSING DATA ----
DataExplorer::plot_missing(df)+
  theme_bw()+
  theme(legend.position = "bottom")+
  labs(title="Percentage of Missing per column")
#### GRAPHICAL EXPLORATION OF DATA ----
DataExplorer::introduce(df)

#### META-ANALYSIS ----
test<-df%>%
  filter(Risk_Metric !="R2")%>%
  select(Study_Author_Year, Risk_Estimate, Risk_Estimate_Low, 
         Risk_Estimate_High, Study_Year, Total_N, Study_Design, Risk_Metric)%>%
  mutate(vi = ((Risk_Estimate_High-Risk_Estimate_Low)/3.92)^2,
         yi = Risk_Estimate, 
         Study_Year = as.numeric(Study_Year))%>%
  select(Study_Author_Year, yi, vi, Study_Year, Study_Design, 
         Risk_Estimate, Risk_Estimate_Low, Risk_Estimate_High,
         Total_N,Risk_Metric)%>%
  arrange(Study_Author_Year)
res_meta__test_REMLDL<-meta::metagen(TE = Risk_Estimate,
                           lower=Risk_Estimate_Low,
                           upper=Risk_Estimate_High, 
                           studlab = Study_Author_Year,
                           data = test,
                           sm = "OR",
                           transf=FALSE,
                           fixed = FALSE,
                           random = TRUE,
                           adhoc.hakn.ci ="ci",
                           method.predict = "HK",
                           adhoc.hakn.pi = "se",
                           tau.common=FALSE,
                           method.tau = "REML",
                           title = "Glyphosate and Parkinson")
summary(res_meta__test_REMLDL)
meta::forest(res_meta__test_REMLDL, 
             sortvar = TE, prediction = TRUE, 
             print.tau2 = TRUE,layout = "RevMan5")

meta_df<-df%>%
  filter(Risk_Metric =="OR")%>%
  select(Study_Author_Year, Risk_Estimate, Risk_Estimate_Low, 
         Risk_Estimate_High, Study_Year, Total_N, Study_Design, Risk_Metric)%>%
  mutate(vi = ((Risk_Estimate_High-Risk_Estimate_Low)/3.92)^2,
         yi = Risk_Estimate, 
         Study_Year = as.numeric(Study_Year))%>%
  select(Study_Author_Year, yi, vi, Study_Year, Study_Design, 
         Risk_Estimate, Risk_Estimate_Low, Risk_Estimate_High,
         Total_N, Risk_Metric)%>%
  arrange(Study_Author_Year)
wald_test<-metafor::conv.wald(out=Risk_Estimate, 
                        ci.lb=Risk_Estimate_Low, 
                        ci.ub=Risk_Estimate_High, 
                        data=test,
                        n=Total_N,
                        transf = "log", 
                        check=TRUE)
summary(wald_test)
meta_df%>%filter(Study_Author_Year=="Caballero,2018")%>%
  select(Study_Author_Year, yi,vi)
wald_test%>%filter(Study_Author_Year=="Caballero,2018")%>%
  select(Study_Author_Year, yi,vi)

res<-rma(yi=yi, vi=vi, data=meta_df)
summary(res)
res<-rma(yi=yi, vi=vi, data=meta_df, method="REML")
summary(res)
res<-rma(yi=yi, vi=vi, data=meta_df, method="DL")
summary(res)
res<-rma(yi=yi, vi=vi, data=meta_df, method="EB")
summary(res)
res<-rma(yi=yi, vi=vi, data=meta_df, method="HSk")
summary(res)

df%>%
  select(Study_Author_Year, Risk_Estimate, Risk_Estimate_Low, 
         Risk_Estimate_High)%>%
  mutate(left = abs(Risk_Estimate-Risk_Estimate_Low), 
         right= abs(Risk_Estimate-Risk_Estimate_High), 
         dist = abs(right-left),
         equal_dist = if_else(left == right, 'yes', 'no'))%>%
  select(equal_dist)%>%
  group_by(equal_dist)%>%
  count()

df%>%
  select(Study_Author_Year, Risk_Estimate, Risk_Estimate_Low, 
         Risk_Estimate_High, Risk_Estimate_Difference)%>%
  mutate(left = abs(Risk_Estimate-Risk_Estimate_Low), 
         right= abs(Risk_Estimate-Risk_Estimate_High), 
         dist = abs(right-left),
         equal_dist = if_else(left == right, 'yes', 'no'))%>%
  ggplot()+
  geom_point(aes(x=left, 
                 y=right, 
                 size=Risk_Estimate), alpha=0.5)+
  geom_abline(intercept=0, slope=1, lty=2, col="red")+
  theme_bw()+
  theme(legend.position = "bottom")+
  labs(x="Left distance 95% confidence interval",
       y="Right distance 95% confidence interval", 
       size="Risk estimate", 
       title="Relationship between left and right distance of the 95% confidence interval and the risk estimate", 
       subtitle = "for the review of Ahmed,2017 focused on pesticides")


meta::metagen(TE = yi,
              seTE = sqrt(test$vi)/sqrt(test$Total_N),
              studlab = Study_Author_Year,
              data = meta_df,
              sm = "OR",
              transf=FALSE,
              fixed = FALSE,
              random = TRUE,
              method.tau = "DL",
              title = "Glyphosate and Parkinson")

meta::metagen(TE = Risk_Estimate,
              lower=Risk_Estimate_Low,
              upper=Risk_Estimate_High, 
              studlab = Study_Author_Year,
              data = meta_df,
              sm = "OR",
              transf=FALSE,
              fixed = FALSE,
              random = TRUE,
              method.tau = "REML",
              title = "Glyphosate and Parkinson")
meta::metagen(TE = Risk_Estimate,
              lower=Risk_Estimate_Low,
              upper=Risk_Estimate_High, 
              studlab = Study_Author_Year,
              data = meta_df,
              sm = "OR",
              transf=FALSE,
              fixed = FALSE,
              random = TRUE,
              method.tau = "DL",
              title = "Glyphosate and Parkinson")
meta::metagen(TE = Risk_Estimate,
              lower=Risk_Estimate_Low,
              upper=Risk_Estimate_High, 
              studlab = Study_Author_Year,
              data = meta_df,
              sm = "OR",
              transf=FALSE,
              fixed = FALSE,
              random = TRUE,
              method.tau = "EB",
              title = "Glyphosate and Parkinson")
meta::metagen(TE = Risk_Estimate,
              lower=Risk_Estimate_Low,
              upper=Risk_Estimate_High, 
              studlab = Study_Author_Year,
              data = meta_df,
              sm = "OR",
              transf=FALSE,
              fixed = FALSE,
              random = TRUE,
              method.tau = "HS",
              title = "Glyphosate and Parkinson")

res_meta_DL<-meta::metagen(TE = Risk_Estimate,
                           lower=Risk_Estimate_Low,
                           upper=Risk_Estimate_High, 
                           studlab = Study_Author_Year,
                           data = meta_df,
                           sm = "OR",
                           transf=FALSE,
                           fixed = FALSE,
                           random = TRUE,
                           adhoc.hakn.ci ="ci",
                           method.predict = "HK",
                           adhoc.hakn.pi = "se",
                           tau.common=FALSE,
                           method.tau = "DL",
                           title = "Glyphosate and Parkinson")
summary(res_meta_test_DL)
meta::forest(res_meta_test_DL, 
             sortvar = TE,
             prediction = TRUE, 
             print.tau2 = TRUE,layout = "RevMan")

res_meta_REML<-meta::metagen(TE = Risk_Estimate,
                            lower=Risk_Estimate_Low,
                            upper=Risk_Estimate_High, 
                            studlab = Study_Author_Year,
                            data = meta_df,
                            sm = "OR",
                            transf=FALSE,
                            fixed = FALSE,
                            random = TRUE,
                            adhoc.hakn.ci ="ci",
                            method.predict = "HK",
                            adhoc.hakn.pi = "se",
                            prediction = TRUE,
                            tau.common=FALSE,
                            method.tau = "REML",
                            title = "Glyphosate and Parkinson")
summary(res_meta_REML)
meta::forest(res_meta_REML, 
             sortvar = TE, prediction = TRUE, 
             print.tau2 = TRUE,layout = "RevMan5")


Hmisc::describe(meta_df$yi)
Hmisc::describe(meta_df$vi)


res_meta<-meta::metagen(TE = Risk_Estimate,
                        lower=Risk_Estimate_Low,
                        upper=Risk_Estimate_High, 
                        studlab = Study_Author_Year,
                        data = meta_df,
                        sm = "OR",
                        transf=FALSE,
                        fixed = FALSE,
                        random = TRUE,
                        adhoc.hakn.ci ="ci",
                        method.predict = "HK",
                        adhoc.hakn.pi = "se",
                        prediction=TRUE,
                        tau.common=FALSE,
                        method.tau = "REML",
                        title = "Glyphosate and Parkinson")
summary(res_meta)
res_metafor<-rma(yi=res_meta$TE, 
            sei = res_meta$seTE, 
            method=res_meta$method.tau, 
            slab=res_meta$studlab,
            measure="OR", 
            backtransf="exp")
summary(res_metafor)


permres<-permutest(res_metafor, exact=TRUE, permci = TRUE)
permres
plot(permres, xlim=c(-5,5))
confint(res_metafor, type="PL")

sav <- baujat(res_metafor, symbol=19, xlim=c(0,55))
sav <- sav[sav$x >= 10 | sav$y >= 0.10,]
text(sav$x, sav$y, sav$slab, pos=1, cex=0.8)

inf <- influence(res_metafor)
par(mfrow=c(8,1))
plot(inf)



res_meta_REML_egger<-eggers.test(res_meta_REML)
summary(res_meta_REML_egger)
plot(res_meta_REML_egger)
regtest(res_metafor)
funnel(res_metafor, refline=1.24)
funnel(res_metafor, level=c(90, 95, 99), 
       shade=c("white", "gray55", "gray75"), refline=1.24, legend=TRUE)

res_meta_trimfill<-meta::trimfill(res_metafor, 
                                  )
summary(res_meta_trimfill)
plot(res_meta_trimfill)

taf <- trimfill(res_metafor, estimator="R0")
taf
funnel(taf, legend=list(show="cis"))

taf <- trimfill(res_metafor)
filled <- data.frame(yi = taf$yi, vi = taf$vi, fill = taf$fill)
rma(yi, vi, data=filled)


res_meta_test_REML_design<-update(res_meta_REML, 
                                  subgroup=Study_Design, 
                                  tau.common=FALSE, prediction=TRUE, 
                                  prediction.subgroup=TRUE)
summary(res_meta_test_REML_design)
meta::forest(res_meta_test_REML_design, 
             sortvar = TE,
             prediction = TRUE, 
             print.tau2 = TRUE,layout = "RevMan5")


meta_metric_df<-df%>%
  select(Study_Author_Year, Risk_Estimate, Risk_Estimate_Low, 
         Risk_Estimate_High, Study_Year, Total_N, Study_Design, Risk_Metric)%>%
  mutate(vi = ((Risk_Estimate_High-Risk_Estimate_Low)/3.92)^2,
         yi = Risk_Estimate, 
         Study_Year = as.numeric(Study_Year))%>%
  select(Study_Author_Year, yi, vi, Study_Year, Study_Design, 
         Risk_Estimate, Risk_Estimate_Low, Risk_Estimate_High,
         Total_N, Risk_Metric)%>%
  arrange(Study_Author_Year)
res_meta_metric_REML<-meta::metagen(TE = Risk_Estimate,
                             lower=Risk_Estimate_Low,
                             upper=Risk_Estimate_High, 
                             studlab = Study_Author_Year,
                             data = meta_metric_df,
                             transf=FALSE,
                             fixed = FALSE,
                             random = TRUE,
                             adhoc.hakn.ci ="ci",
                             method.predict = "HK",
                             adhoc.hakn.pi = "se",
                             prediction = TRUE,
                             tau.common=FALSE,
                             method.tau = "REML",
                             title = "Glyphosate and Parkinson")
summary(res_meta_metric_REML)
res_meta_test_REML_metric<-update(res_meta_metric_REML, 
                                  subgroup=Risk_Metric, 
                                  tau.common=FALSE, prediction=TRUE, 
                                  prediction.subgroup=TRUE)
summary(res_meta_test_REML_metric)
meta::forest(res_meta_test_REML_metric, 
             sortvar = TE,
             prediction = TRUE, 
             print.tau2 = TRUE,layout = "RevMan5")




