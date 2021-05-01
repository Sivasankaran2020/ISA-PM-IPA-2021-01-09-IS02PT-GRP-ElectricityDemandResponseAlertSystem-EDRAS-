### Electricity Demand Response Alert System (EDRAS)

---

## SECTION 1 : PROJECT TITLE
## Electricity Demand Response Alert System (EDRAS) - AI based solution predicting Electricity Price & Demand.

---

## SECTION 2 : EXECUTIVE SUMMARY / PAPER ABSTRACT

**Summary**

With the increase of both electrical demand and renewable electrical resources, the demand response program will provide a major source of reliable and flexible capacity to ensure resource adequacy cost-effectively. The Forecast and notification system that the team implemented will enable potential Demand Response Market participants to have a pre-emptive notification and bid into the system during the forecasted periods that will have a high electrical USEP price. The eventual market benefit will enable the market participant to have a revenue of around $4500/MWhr. ( Taking an example of a chiller plant which has a dispatchable load of 10MWhr, this will translate to a potential revenue of $45K per hour of participation when triggered. 

**Objective**

Real-Time Notification of High USEP for Demand Response Participation To participate in the programme, licensed load providers will have to submit demand bids into the NEMS for the periods in which they intend to offer load curtailments through a process called demand side bidding. The demand bids for a particular trading period will consist of the following: total load (in MW) for the registered block, a series of energy curtailment and price quantity tranches and the linear ramp rates (MW/min) as determined during the registration stage. 
Submission of the bids are required to be completed within 65min ( “Gate Closure”) of the trading, hence active real-time notification will allow the participant to have a “look-ahead” plan in order to bid into the system during the periods where the electricity prices are foreseen to be high.


**Development Methodology**

The approach follows the CRISP-DM process. It consists of five stages, as listed below:
1)	Understand the Business Requirements
2)	Understanding & Analyzing the Data
3)	Verify / Optimize the model
4)	Analyze the results and reveal the insights


**Technologies Used**

The model developed for the project work uses Python and following open source libraries:

•	Microsoft Power Query:  this component will get the data from EMA website to be shown in the Power BI

•	Microsoft Power BI: Power BI is used to display the data and show the 5 price points to trigger the alert whenever USEP price has reached certain threshold.

•	Microsoft Power Automate: this component will drive the workflow.

•	Email (Microsoft Outlook): one of the alert format for the user is in the form of email

•	Telegram: another alert will be in the format of telegram message to ensure that users will get notification in their mobile phone

•	Python: To scrap the latest EMC data with TagUI (RPA) from the website and append to the training dataset and predict the Demand and USEP with ML model

---

## SECTION 3 : CREDITS / PROJECT CONTRIBUTION

| Official Full Name  | Student ID (MTech Applicable)  | Email (Optional) |
| :------------ |:---------------:|  :-----|
| Harry Chan | A0213530X | e0508631@u.nus.edu |
| Chong Keng Han | A0213547H | e0508648@u.nus.edu |
| Wen Cheng | A0213572L | e0508673@u.nus.edu |
| Sivasankaran Balakrishnan | A0065970X | e0507972@u.nus.edu |

---

## SECTION 4 : VIDEO OF SYSTEM MODELLING & USE CASE DEMO
<div align="center">
  <a href="https://youtu.be/LX3SzWYeCGo">
    <img src="http://i3.ytimg.com/vi/LX3SzWYeCGo/hqdefault.jpg" alt="CHURN Project Video">
  </a>
</div>

---

## SECTION 5 : INSTALLATION & USER GUIDE

`Refer to "Electricity Demand Response Alert System (EDRAS) - Project Insatallation-User Guide.docx" document at Github Folder: ProjectReport`

---

## SECTION 6 : PROJECT REPORT / PAPER

`Refer to "Electricity Demand Response Alert System (EDRAS) - Project Report.docx' document at Github Folder: ProjectReport`

---

## SECTION 7 : CONCLUSION

With the increase of both electrical demand and renewable electrical resources, the demand response program will provide a major source of reliable and flexible capacity to ensure resource adequacy cost-effectively. The Forecast and notification system that the team implemented will enable potential Demand Response Market participants to have a pre-emptive notification and bid into the system during the forecasted periods that will have a high electrical USEP price. The eventual market benefit will enable the market participant to have a revenue of around $4500/MWhr. ( Taking an example of a chiller plant which has a dispatchable load of 10MWhr, this will translate to a potential revenue of $45K per hour of participation when triggered. 


## SECTION 
---

