# VBA Method for slope residuals for multivariate time series in biology

### About experiment
Suppose we have been collecting multivariate data over a period of time. Each of property behaves independently lineary with respect to time. At some point we decided to influence on the whole experiment by changing the slopes. At that point each property change slopes differently.
*We are interested in the estimate of the slope for each property separately, between the residuals from the second part of the experiment and the estimate of the slope of the first part of the experiment.*

### Data description

Time should be in all odd columns, a data should be in even columns.

| Time|Data1| Time|Data2| Time|Data3|
| :-: | :-: | :-: | :-: | :-: | :-: |
| 0,40|740,5| 0,40|511,7| 0,40|922,5|
|10,40|737,3|10,40|508,8|10,40|924,1|
|20,40|728,9|20,40|502,4|20,40|919,1|
|30,40|725,8|30,40|500,2|30,40|920,2|
|40,40|720,0|40,40|496,0|40,40|916,5|
|50,40|719,2|50,40|494,7|50,40|918,3|
|60,40|711,7|60,40|489,5|60,40|909,6|
|70,40|710,5|70,40|488,1|70,40|910,3|
|80,40|705,7|80,40|484,0|80,40|905,7|
|90,40|702,0|90,40|481,3|90,40|903,0|

### About algorithm
Algorithm uses RANSAC (https://en.wikipedia.org/wiki/Random_sample_consensus) to remove outliers, and compute slopes on inliers. Number of picked inliers will be defined as percentage. The parameter is named: *percentageOfInliers*. 

Also, by default it will run 20 iterations, which proved to be sufficient in experiments. 

### How to use

#### Running

1. Enable developer mode in excel: https://support.microsoft.com/en-au/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45
2. In Developer tab > Visual basic > View > Code
3. Copy code to Code editor
4. Select only Numbered cells
5. Go to Macro > select ComputeSLOPEs > Run

Under each column should be the result.

#### Modifying script

There are 4 main variables: *beginIndex*, *middleIndex*, *endIndex*, *deviateFromMiddle*.

These variables are used to help algorithm to compute slopes. As mentioned above in **About experiment**, there are two parts in each time series.
- *beginIndex* - guess of start of series. Everything before this index will not be included.
- *middleIndex* - guess of the point in time where we influenced the experiment.
- *endIndex* - guess of the ending point. Everything after this index will not be included.
- *deviateFromMiddle* - deviation from *middleIndex*. Helps to separate the first part from the second.

So algorithm will be computed on two segments: [*beginIndex*, *middleIndex* - *deviateFromMiddle*] and [*middleIndex* + *deviateFromMiddle*, *endIndex*].


