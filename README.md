# VBA_Cascading-Text-Classifier
This is a custom ML algorithm I developed and implemented in VBA to classify text data.

## 1.0 Introduction
### 1.1 Original Project
Along with working with data, one of my passions is personal finance so since 2016 I have tracked and classified every single transaction I've made (groceries, rent, restaurants, transfer to/from specific accounts, etc.).
This allows me to see how my money flows and grows, how my spending compares to budget, and gives me great data for future financial planning.

First I developed a flexible VBA based ETL process to load downloaded data from my various banking institutions into my Finance Dashboard Excel sheet.
Next I developed this machine learning algorithm and codebase to automatically classify transactions as a final step to the loading process.

Why VBA? ... well, it was my strongest programming language at the time, I already had a robust framework developed in Excel for the data, and I thought it would be fun.

### 1.2 Problem Description
My goals were to:
1. Correctly classify transactions based on previous historical transaction descriptions

2. Not guess a classification if certainty is too low

3. Be able to classify completely new transaction descriptions based on similarly named transactions
    - I.e. if I had already been to `Bob's No Frills Hamilton` and classified it as `Groceries`, it should be able to tell that `Dave's No Frills Burlington` is also `Groceries` despite never having been there before.

4. Be able to classify identical transaction descriptions differently depending on the source account of the data
    - At some banks if you transfer money between your chequing/savings, the transaction descriptions may only say `Transfer In` / `Transfer Out`.  Therefore if the transaction occurs from your chequing account, the algorithm should know those transactions could only come from your savings account (since you can’t transfer money from one account to the exact same account).
   
5. Classify a monthly batch of transactions (typically 10-100) in <5  seconds (this is fast enough for my personal use)

6. Generalize the code enough that the algorithm could be used for different applications


## 2.0 Detailed Explanation
### 2.1 Classifier Algorithm
**Classification Call Stack:**

`GuessValue`<-`CascadeGuessValue`<-*Custom_Implementation_function*

#### 2.1.1 GuessValue
At the lowest level (`GuessValue`), the algorithm simply calculates the probability that a string belongs to each class and then picks the highest. **This satisfied requirement 1.**

![P(class)=\frac{\sum[{class\. matches}]}{\sum[{matches}]}](https://latex.codecogs.com/gif.latex?\bg_white&space;P(class)=\frac{\sum[{class\.&space;matches}]}{\sum[{matches}]})

#### 2.1.2 CascadeGuessValue
At the next level up (`CascadeGuessValue`), the algorithm takes the input string and runs `GuessValue`, then evaluates the result.  If the probability from `GuessValue` is at or above a certain tolerance and a minimum number of matches has been found, then the prediction is accepted. **This satisfied requirement 2.**

Otherwise the prediction is thrown away and the input string is "cascaded" into more substrings.
Cascading creates substrings by splitting the input string using space (" ") as a delimiter into pieces, and then joining these pieces together (preserving string order) to create the desired number of substrings.
    
    E.g.
        CascadeString("Hello world foo bar")
        Substrings = 1 -> "Hello world foo bar"

        Substrings = 2 -> "Hello world foo",
                          "world foo bar"

        Substrings = 3 -> "Hello world",
                          "world foo",
                          "foo bar"

        Substrings = 4 -> "Hello",
                          "world",
                          "foo",
                          "bar"

`GuessValue` is then run on the new set of substrings, and again returns a result if found acceptable, otherwise it will cascade the input string to the next level. **Cascading was implemented to satisfy requirement 3.**

Note, `CascadeGuessValue` could be modified to use a different evaluation function than `GuessValue`.

### 2.1.3 Notes on Model Persistence and Scaling
This code doesn't create a persistent model.
It "retrains" each time it is implemented.  This is efficient enough for my purposes and scale.

However if this needed to be drastically scaled up, the model could be made persistent by storing a 2D "result" array of probability scores for each string and substring on one axis, and each class on the other axis. This would be the model fitting/training procedure.

To implement further dimensionality (changing the probability score based on a feature other than the input string, e.g. for my use case, considering the account where a transaction occurred), more dimensions could be added the "result" array.

Then predicting with the trained model would just be a matter of looking up from the "result" array: the class with the highest probability score for the given substring and another other given dimensions .

## 2.2 Dataset Filtering/Boolean Masking
Providing a filtered training data set can simulate adding more features to model training.

A boolean masking framework was developed to implement this.

**Data Filter Call Stack:**

*Custom_Filter_function*<-`GenerateFilter`<-*Custom_Implementation_function*

The `GenerateFilter` function will generate a boolean mask that can be passed to `CascadeGuessValue` to tell it to skip over rows of data when doing an evaluation. **This was implemented to satisfy requirement 4.**

Since VBA's array functionality is limited, I also create the `LogicArray` function to allow for logical operations (AND, OR, NOT) on boolean masks.

## 3.0 Implementation
### 3.1 Transaction Categorization
The `GuessCategory` function shows how I implemented the classifier along with the boolean masking functionality to classify my transaction data. I was able to satisfy requirement 4 by first filtering the training data set to only include transactions from the same account.  If no classification was found on this first pass, then the filter is removed to increase the likelihood of matches being found.

### 3.2 Transaction Filename - Account Matching
The `GuessAccount` function is a different implementation of the classifier code that I used to slightly streamline my ETL process.  The ETL process is to:
1. Download account transactions to `.csv`s
2. Click the `import data` button in my Dashboard file
3. From a userform, select the file to import
4. From a userform, select the bank account to which the transactions belong
5. Code transforms data to fit into transaction table, and auto-classifies transactions

`GuessAccount` was used to determine which account is likely being imported based on the `.csv` file name, and pre-highlight that account option in the userform for step 4 of the ETL.  It is a very small quality of life improvement, but **demonstrates that the code satisfied requirement 6.**

## 4.0 Classification Performance
TO DO
- ROC Curve, tolerance tuning
- Classification Time

## 5.0 Dependencies
https://github.com/dnissim/VBA_General

A general library of functions for interacting with different data objects in Excel, and effectively working with array data in memory.