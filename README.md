# Stock-Analysis with VBA
### Challenge 2: VBA of Wall Street
## Project Background
My friend Steve asked for assitance analyzing the stock performace history for DAQO New Energy Corp (DQ) and other green energy stocks. His parents invested solely in DQ and he wants to ensure they made a viable investment. I put together a workbook to analyze the data he provided for DQ and the other green energy stocks for 2017 and 2018 but he wants to expand the research. He now wants to include the entire stock market over the last few years. However, my previous code may not work well for all of the stocks he wants to analyze so I have refactored the code to sustain more data. Now, he will be able to analyze more than the 12 stocks originally included in his data set.
THe analysis below explores the difference in run times between the original and refactored codes, the pros and cons of both, and provides a brief analysis of the stock returns.

## Findings
Before refactoring the data set the code ran for 0.9297 seconds for the 2017 data and for 0.9065 seconds for the 2018 data. After refactoring the data, the time from start to finished increased substantially across both years. As you can see in the images below, both years now take over 7 seconds to run.
<img width="277" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/96352625/149645053-6169e531-4bbb-4b88-b68a-ec6736510c64.png">
<img width="277" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/96352625/149645056-94c99c00-927f-4903-960b-4be1f6769733.png">

As a check on the accuracy of my code in the refactored analysis, both codes derived the same results. As we can see, all green energy stocks performed better in 2017 than the following year where a strong majority of the stocks performed negative returns. DQ, the stock of interest to Steve and his parents, saw returns of over 199% in 2017 but ran in the red in 2018 at -62.6%. I recommend Steve take these findings to his parents to explore investing in alternative stocks and diversifying their portfoloio. However, we only have access to '17 & '18 data. Subsequent years may see different trends.
<img width="292" alt="VBA_Challenge_2017_Table" src="https://user-images.githubusercontent.com/96352625/149645194-f876afc6-99e8-4f3d-8984-1d7cfdc6438d.png">
<img width="286" alt="VBA_Challenge_2018_Table" src="https://user-images.githubusercontent.com/96352625/149645197-551bc127-5982-4f8f-b3cc-652ce6b43023.png">

## Summary

