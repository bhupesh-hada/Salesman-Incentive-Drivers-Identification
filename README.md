# Objective
The aim of the project is to identify the best Key Performance Inidicators (KPIs) or features and recommend their optimum values to best incentivize a salesman. Separate KPIs recommenations are made for different clusters of salesman. These cluster differ w.r.t to the salesman route type.

# Methodology
There are more than 100 features extracted and collated from 6 data sources. EDA was performend to identify top 20-25 most important features. Different combinations of these features are used while model training as the inclusion of some features is conditioned on the presence of other features. On each of these combinations stepwise regression is performed to get the best feature.

# Results
For a given cluster best model is selected based on adjusted rsquare and f_statisitcs. In the final model, features with p-value < 0.001 and positive coefficient are classified as primary recommendation, p-value < 0.05 and postive coefficients as secondary recommendation and the remaining as not recommended
![SIO_2](https://user-images.githubusercontent.com/24865203/189040519-7755072b-1cb8-40f9-b52d-abae45e552fc.PNG)

Finally, for the recommended KPIs maximum or minimum levels are also presented to help facilitate th decision on the allocation of the incentive budget
* The blue curve from left to right shows the increment in % Achievement of Target GSV (depedent variable) as one pushes to increase performance in this KPI / Measure. This allows comparison on how much impact can be achieved by increasing the different recommended KPIs/Measure above the current average level.
* The distance between the average level and the minimum / maximum level help provide a view on how much further one should push to increase the different KPIs / Measures.
![SIO_3](https://user-images.githubusercontent.com/24865203/189041254-72eb46ea-9947-48d1-bafe-49bc35e48056.PNG)
![SIO_4](https://user-images.githubusercontent.com/24865203/189041332-aee43715-fe28-4e04-9676-e7efab6fd0ec.PNG)


