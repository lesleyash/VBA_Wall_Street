# VBA_Wall_Street
Assignment for VBA_Wall_Street

## Overview of Project
### Purpose
The purpose of the analysis is to help Steve analyze the performance of stock in terms of return and trading volumes. 
We will also refactor the original code to reduce the execution time of the VBA module

## Results

### Stock Performance
 - Based on the stock performance results shown in the table below, the selected stocks generated a much better return in year 2017 than 2018.
 - Stocks ticker **RUN** and **ENPH** are the best performers in year 2018 
 - Between Stock **RUN** and **ENPH**, Stock **ENPH** shows consistent strong performance in both year 2017 and year 2018

#### Stock performance of year 2017
![image](https://user-images.githubusercontent.com/92648619/141701982-e5892b2d-97da-4676-85d4-33f6e630adb2.png)

#### Stock performance of year 2018
![image](https://user-images.githubusercontent.com/92648619/141702000-5575af9c-c16d-43f7-86bd-d19b52fa99be.png)

### Comparison on Execution time
The original code requires ~0.5 seconds to run the stock performance each year, while the refactored code only require ~0.1 second.
The main difference between the original code and refactored code are the number of loops it need to run in order to get results. The refactored codes only need to run one loop when the original code need to run 12 loops, which means one loop per each stock.

#### Execution time of original code
![image](https://user-images.githubusercontent.com/92648619/141702367-f7499d6b-e557-4639-a077-1de86adf6850.png)
![image](https://user-images.githubusercontent.com/92648619/141702379-efb31e42-03cf-422c-b0fa-6d447db93b34.png)

#### Execution time of refactored code
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92648619/141702269-f7e6a58e-e1fb-4fa9-a506-c04c83d1861a.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/92648619/141702259-7934e1bc-aec6-4f64-ab7f-1a98874d2e60.png)


## Summary

### Advantage and Disadvantage of refactoring code
#### Advantage
The advantage of refactoring is to shortern the time the module needs to run the code, which helps expand the data the code can be run.
And refactoring usually means simplier or shorter code module, which can help people understand and maintain the code.

#### Disadvantage
The disadvantage is the amount of time and tests you have to put in the refactoring. If the results are not significantly better, the return of the time you invest in refacoring code might be very small or even negative.

### Pros and Cons of original and refactoring
#### Pros and Cons of original code
**Pros** of the original code is there is no need to create array for volume and price. 
It does not require the order of the original data to be exactly same as the ticker order we listed on code since it will run the entire loop for every ticker.
![image](https://user-images.githubusercontent.com/92648619/141702894-1516e015-4100-4616-a5cc-321f0b08f1f7.png)

**Cons** of the original code is the amount of time the code requires to run due to 12 complete loops needed.

#### Pros and Cons of refactoring code
***Pros*** of the refactoring code is the significant reduction of code executionn time.
Also, the refactored code only require one For Loop statement, which is easiler to follow and maintain.

***Cons*** of the refactoring code is any change in the order of the ticker in the original data sheet may results errors due to logic of the refactored code.


