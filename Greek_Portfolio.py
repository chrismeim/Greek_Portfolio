import pandas as pd
import numpy as np

'''
Import my Dataset from my excel file
'''

Hellenic_Petroleum_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "Hellenic Petroleum")
MotorOil_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "MotorOil")
OPAP_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "OPAP")
Terna_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx",sheet_name = "Terna" )
Jumbo_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx",sheet_name = "Jumbo")
Quest_Holdings_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "Quest Holdings")
Alpha_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "Alpha Bank")
Mytilineos_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx",sheet_name = "Mytilineos")
Aegean_index = pd.read_excel(r"C:\Users\Chris\Desktop\Greek_Stocks_Project.xlsx", sheet_name = "Aegean")

def data_manipulator(dataset, name):  #takes the dataset of a company and changes the format of the data in order to subtract the dates and adj_close_price
    list1 = []
    for i in range(len(dataset)):
        j = 0
        x = dataset.iloc[i][0]   
        word = ""
        sentence = []
        while j < len(x):
            if x[j] != ",":
                word += x[j]
                j +=1
            else:
                sentence.append(word)
                word = ""
                j +=1
        list1.append(sentence)
    
    date = []
    adj_close = []
    for i in range(len(dataset)):
        date.append(list1[i][0])
        adj_close.append(list1[i][4])

    name = {"Date":date, "Adj_Close":adj_close}
    name = pd.DataFrame(name)    
    return name


Hellenic_Petroleum = {}
MotorOil = {}
OPAP = {}
Terna = {}
Jumbo = {}
QuestH = {}
Alpha = {}
NationalBank = {}
Mytilineos = {}
OTE = {}
Aegean = {}

Hellenic_Petroleum = data_manipulator(Hellenic_Petroleum_index, Hellenic_Petroleum)
MotorOil = data_manipulator(MotorOil_index, MotorOil)
OPAP = data_manipulator(OPAP_index, OPAP)
Terna = data_manipulator(Terna_index, Terna)
Jumbo = data_manipulator(Jumbo_index, Jumbo)
QuestH = data_manipulator(Quest_Holdings_index, QuestH)
Alpha = data_manipulator(Alpha_index, Alpha)
Mytilineos = data_manipulator(Mytilineos_index, Mytilineos)
Aegean = data_manipulator(Aegean_index, Aegean)

Aegean["Adj_Close"] = Aegean["Adj_Close"].apply(lambda x: float(x.strip('""')))
MotorOil["Adj_Close"] = MotorOil["Adj_Close"].apply(lambda x: float(x.strip('""')))
QuestH["Adj_Close"] = QuestH["Adj_Close"].apply(lambda x: float(x.strip('""')))
Jumbo["Adj_Close"] = Jumbo["Adj_Close"].apply(lambda x: float(x.strip('""')))
Terna["Adj_Close"] = Terna["Adj_Close"].apply(lambda x: float(x.strip('""')))
Alpha["Adj_Close"] = Alpha["Adj_Close"].apply(lambda x: float(x.strip('""')))
Mytilineos["Adj_Close"] = Mytilineos["Adj_Close"].apply(lambda x: float(x.strip('""')))
OPAP["Adj_Close"] = OPAP["Adj_Close"].apply(lambda x: float(x.strip('""')))
Hellenic_Petroleum["Adj_Close"] = Hellenic_Petroleum["Adj_Close"].apply(lambda x: float(x.strip('""')))


list_of_companies = [Hellenic_Petroleum, MotorOil, OPAP, Terna, Jumbo, QuestH, Alpha, Mytilineos, Aegean]
names_of_companies = ["Hellenic_Petroleum", "MotorOil", "OPAP", "Terna", "Jumbo", "QuestH", "Alpha", "Mytilineos", "Aegean"]


'''
Here I create this new data frame called companies, where I will list the closing prices of all the companies of my portfolio
'''

companies = {}
companies["Date"] = OPAP["Date"]
companies["Aegean"] = Aegean["Adj_Close"]
companies["MotorOil"] = MotorOil["Adj_Close"]
companies["QuestH"] = QuestH["Adj_Close"]
companies["Jumbo"] = Jumbo["Adj_Close"]
companies["Terna"] = Terna["Adj_Close"]
companies["Alpha"] = Alpha["Adj_Close"]
companies["Mytilineos"] = Mytilineos["Adj_Close"]
companies["OPAP"] = OPAP["Adj_Close"]
companies["Hellenic_Petroleum"] = Hellenic_Petroleum["Adj_Close"]

companies = pd.DataFrame(companies)

companies["Date"] = pd.to_datetime(companies["Date"])
companies.set_index("Date", inplace = True)  #Setting the index to "Date"

import matplotlib.pyplot as plt
companies.plot() #Ploting the data set



'''
Creat a new data frame called returns whre I will have the returns from all the companies of my portfolio
'''

returns = np.log(companies.shift(1)/companies)
returns = returns.fillna(returns.mean())
returns = returns.sort_index(ascending=True) 

returns.mean()*252
returns.std()*252
returns.cov()

'''
I will now try to optimize the returns of my Portfolio by Maximizing the Equation E(Rp) - 1/2*gamma*Var(Rp)
E(Rp) = Expected Return of Portfolio
Var(Rp) = Volatility of Portfolio
gamma = risk aversion factor
'''

cov_matrix = returns.cov()
gamma = 5   #Gamma gives us how risk averse we are, using a gamma of 5 we assume that we are moderate investors
risk_free_rate = 0.039    #19/11/2023 Greek 10yr government bond yield

inv_cov_matrix = np.linalg.inv(cov_matrix) #this gives us the inverse of the covariance matrix
excess_returns = returns.mean() - risk_free_rate
weights = (1/gamma) *np.dot(inv_cov_matrix, excess_returns)
weights /= np.sum(weights)


Total_money = int(input("Give me how much money you are going to invest in total "))
Percentage_allocated_in_riskfree = 0.6
Euros_allocated_risky = Total_money*(1-Percentage_allocated_in_riskfree)
Euros_allocated_riskfree = Total_money*Percentage_allocated_in_riskfree

x = returns.mean()*252 #Yearly return
y = returns.std()*(250**0.5) #Yearly volatility
z = weights*Euros_allocated_risky #money for a 4000 Portfolio (10.000 * 0.4 which is the money you are going to allocate in risky assets - stocks)

import math

number_of_shares = []
for i in range(9):
    number_of_shares.append(math.floor(z[i] / companies.iloc[0][i]))

sectors = ["Energy", "Energy","Entertainment", "Construction", "Retail", "Consumer Services", "Banking", "Construction", "Transportation"]

'''
I create a table called Return_Risk where I will display certain data for my risky portfolio such as allocated weights, Expected returns and more...
'''

Return_Risk = {"Sector":sectors,"Expected_Returns":x, "Volatility":y, "Optimal_Weights":weights, "Allocated_Money":z, "Number of Shares":number_of_shares}
Return_Risky = pd.DataFrame(Return_Risk)

print(Return_Risky)




'''
Now I will present some specific characteristics of my portfolio such as overall expected returns and volatility, along with some key facts about our risk free positionn
'''

number_of_bonds = Euros_allocated_riskfree/103  #103 is the current price of the Greek 10yr Bond 19/11/2023
number_of_bonds = math.floor(number_of_bonds)
print(f"The number of Bond that you have to buy is {number_of_bonds}")

Expected_return_of_risky_portfolio = np.dot(x, weights.T)
print(f"The expected return of the risky portfolio is {Expected_return_of_risky_portfolio*100}%")

y1 = returns.cov()*250
Expected_volatility_of_risky_portfolio = np.dot(weights, np.dot(y1, weights.T))**0.5
print(f"The expected volatility of the risky portfolio is {Expected_volatility_of_risky_portfolio*100}%")


Expected_return_of_riskfree_portfolio = 0.039
Expected_volatility_of_riskfree_portfolio = 0 #Assuming no default and that you will hold the bond till maturity

Expected_return_of_overall_portfolio = Percentage_allocated_in_riskfree*0.039 + (1-Percentage_allocated_in_riskfree)*Expected_return_of_risky_portfolio
print( f" The expected return of the overall portfolio is {Expected_return_of_overall_portfolio*100}%")

Expected_volatility_of_overall_portfolio = (1-Percentage_allocated_in_riskfree)*Expected_volatility_of_risky_portfolio
print(f" The expected volatility of the portfolio is {Expected_volatility_of_overall_portfolio*100}%")





















