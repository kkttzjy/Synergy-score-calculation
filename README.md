# Synergy-score-calculation
Calculation.r file includes two R functions -- analyze() and Score(). These two R functions are used to get the final synergy score for a batch of data sets. The final synergy score is similar to Charlice score, where the same algorithm is used to calculate the synergy score (https://www.horizondiscovery.com/media/resources/Miscellaneous/software/HD%20Technical%20Manual%20-%20Chalice%20Analyzer+Viewer%20.pdf).

The first one ("analyze") will prepare templates for Combenefit and Chalice software. It also prepares and saves r data to make plots. Once you get the templates for Combenefit in one folder whose name starts with "Combenefit"", you can set this folder as the file path for Combenefit and run batch analysis. You only need to run LOEWE model to get synergy distribution matrix by saving everything during analysis from Combenefit. To run "analyze", you need to specify the raw data name and data information file name. The data information file contains the plate number, cell line name, drug name and dosage. An example information file is provided as "20181019 resazurin infomation.xlsx", and "20181019 resazurin R.xlsx" is a raw data example.

The second function ("Score") can then be used to calculate synergy scores and generate plots. Those results will be all saved in the same folder whose name starts with "Summary results". To run "Score", you also need to specify the raw data name and data information file name. In the "Summary results XX" folder, there is an excel file that contains all synergy scores and inhibition volumn for each data set. In addition, there is a pdf plot file comparing relative viability between monotherapies and combinations for each data set.
