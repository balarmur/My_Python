import pandas as pd
import datetime

"""
This is the python program has the calculation 
"""

startTime = datetime.datetime.now()
print (startTime)

diff=0;perc=0;grocTot=0;suppTot=0;bulkTot=0;gmTot=0;dairyTot=0;frozenTot=0;meatTot=0 
grocPerc=0;grocInfla=0;suppPerc=0;suppInfla=0;bulkPerc=0;bulkInfla=0;gmPerc=0;gmInfla=0;dairyPerc=0;dairyInfla=0;frozenPerc=0;frozenInfla=0;meatPerc=0;meatInfla=0
whseLst=[];prodidLst=[];keyLst=[];qohP5Lst=[];costP5Lst=[];amountP5Lst=[];qohP12Lst=[];costP12Lst=[];amountP12Lst=[];prodLst=[];cateLst=[];proddescLst=[];diffLst=[];percLst=[]
grocPercLst=[];grocInflaLst=[];suppPercLst=[];suppInflaLst=[];bulkPercLst=[];bulkInflaLst=[];gmPercLst=[];gmInflaLst=[];dairyPercLst=[];dairyInflaLst=[]
frozenPercLst=[];frozenInflaLst=[];meatPercLst=[];meatInflaLst=[]

westUpd1 = pd.read_excel('WestUpd.xlsx')
westUpd2 = pd.read_excel('WestUpd.xlsx')

for row,col in westUpd1.iterrows():
	whseLst.append(col[0])
	prodidLst.append(col[1])
	keyLst.append(col[2])
	qohP5Lst.append(int(col[3]))
	costP5Lst.append(float(col[4]))
	amountP5Lst.append(float(col[5]))
	qohP12Lst.append(int(col[6]))
	costP12Lst.append(float(col[7]))
	amountP12Lst.append(float(col[8]))
	prodLst.append(col[9])
	cateLst.append(col[10])
	proddescLst.append(col[11])

	if col[10] == 'Grocery':
		grocTot = grocTot + float(col[5])
	if col[10] == 'Supplements/Vitamins/Health & Beauty Aids':
		suppTot = suppTot + float(col[5])
	if col[10] == 'Bulk':
		bulkTot = bulkTot + float(col[5])
	if col[10] == 'General Merchandise':
		gmTot = gmTot + float(col[5])
	if col[10] == 'Dairy':
		dairyTot = dairyTot + float(col[5])
	if col[10] == 'Frozen':
		frozenTot = frozenTot + float(col[5])
	if col[10] == 'Meat':
		meatTot = meatTot + float(col[5])

for row,col in westUpd2.iterrows():
	
	diff = float(col[4]) - float(col[7])
	diffLst.append(float(diff))
	
	try:
		perc = ((float(col[4]) - float(col[7]))/float(col[7]))
	except:
		perc = 0
	percLst.append(float(perc))

	if col[10] == 'Grocery':
		try:
			grocPerc = float(col[5]) / float(grocTot)
		except:
			grocPerc = 0
	else:
		grocPerc = 0
	grocPercLst.append(float(grocPerc)) 

	if col[10] == 'Grocery' and (abs(float(grocPerc)*100) > 0.0025):
		grocInfla = float(grocPerc) * float(perc)
	else:
		grocInfla = 0
	grocInflaLst.append(float(grocInfla))

	if col[10] == 'Bulk':
		try:
			bulkPerc = float(col[5]) / float(bulkTot)
		except:
			bulkPerc = 0
	else:
		bulkPerc = 0	
	bulkPercLst.append(float(bulkPerc)) 

	if col[10] == 'Bulk' and (abs(float(bulkPerc)*100) > 0.0025):
		bulkInfla = float(bulkPerc) * float(perc)
	else:
		bulkInfla = 0
	bulkInflaLst.append(float(bulkInfla))

	if col[10] == 'Dairy':
		try:
			dairyPerc = float(col[5]) / float(dairyTot)
		except:
			dairyPerc = 0
	else:
		dairyPerc = 0
	dairyPercLst.append(float(dairyPerc)) 

	if col[10] == 'Dairy' and (abs(float(dairyPerc)*100) > 0.0025):
		dairyInfla = float(dairyPerc) * float(perc)
	else:
		dairyInfla = 0
	dairyInflaLst.append(float(dairyInfla))

	if col[10] == 'Frozen':
		try:
			frozenPerc = float(col[5]) / float(frozenTot)
		except:
			frozenPerc = 0
	else:
		frozenPerc = 0		
	frozenPercLst.append(float(frozenPerc)) 

	if col[10] == 'Frozen' and (abs(float(frozenPerc)*100) > 0.0025):
		frozenInfla = float(frozenPerc) * float(perc)
	else:
		frozenInfla = 0
	frozenInflaLst.append(float(frozenInfla))	

	if col[10] == 'General Merchandise':
		try:
			gmPerc = float(col[5]) / float(gmTot)
		except:
			gmPerc = 0
	else:
		gmPerc = 0
	gmPercLst.append(float(gmPerc)) 

	if col[10] == 'General Merchandise' and (abs(float(gmPerc)*100) > 0.0025):
		gmInfla = float(gmPerc) * float(perc)
	else:
		gmInfla = 0
	gmInflaLst.append(float(gmInfla))	

	if col[10] == 'Supplements/Vitamins/Health & Beauty Aids':
		try:
			suppPerc = float(col[5]) / float(suppTot)
		except:
			suppPerc = 0
	else:
		suppPerc = 0
	suppPercLst.append(float(suppPerc)) 

	if col[10] == 'Supplements/Vitamins/Health & Beauty Aids' and (abs(float(suppPerc)*100) > 0.0025):
		suppInfla = float(suppPerc) * float(perc)
	else:
		suppInfla = 0
	suppInflaLst.append(float(suppInfla))	

	if col[10] == 'Meat':
		try:
			meatPerc = float(col[5]) / float(meatTot)
		except:
			meatPerc = 0
	else:
		meatPerc = 0
	meatPercLst.append(float(meatPerc)) 

	if col[10] == 'Meat' and (abs(float(meatPerc)*100) > 0.0025):
		meatInfla = float(meatPerc) * float(perc)
	else:
		meatInfla = 0
	meatInflaLst.append(float(meatInfla))	

dict = {'Warehouse':whseLst,'Product ID':prodidLst, 'Key':keyLst,'P5 QOH':qohP5Lst,'P5 Cost':costP5Lst,'P5 Amount':amountP5Lst,
		'P12 QOH':qohP12Lst,'P12 Cost':costP12Lst,'P12 Amount':amountP12Lst,'Product':prodLst,'Category':cateLst,'Product Desc':proddescLst,
		'Difference':diffLst,'Percentage':percLst,'Grocery':grocPercLst,'Grocery Inflation':grocInflaLst,'Supplements':suppPercLst,'Supplements Inflation':suppInflaLst,
		'Bulk':bulkPercLst,'Bulk Inflation':bulkInflaLst,'General Merchandise':gmPercLst,'GM Inflation':gmInflaLst,'Dairy':dairyPercLst,'Dairy Inflation':dairyInflaLst,
		'Frozen':frozenPercLst,'Frozen Inflation':frozenInflaLst,'Meat':meatPercLst,'Meat Inflation':meatInflaLst}
    
df2 = pd.DataFrame.from_dict(dict)
df2.to_excel('WestUpd03.xlsx',index=False)

print("Grocery Total         : ", grocTot)
print("Grocery Inflation     : ", float(sum(grocInflaLst)))
print("Supplements Total     : ", suppTot)
print("Supplements Inflation : ", float(sum(suppInflaLst)))
print("Bulk Total            : ", bulkTot)
print("Bulk Inflation        : ", float(sum(bulkInflaLst)))
print("GM Total              : ", gmTot)
print("GM Inflation          : ", float(sum(gmInflaLst)))
print("Dairy Total           : ", dairyTot)
print("Dairy Inflation       : ", float(sum(dairyInflaLst)))
print("Frozen Total          : ", frozenTot)
print("Frozen Inflation      : ", float(sum(frozenInflaLst)))
print("Meat Total            : ", meatTot)
print("Meat Inflation        : ", float(sum(meatInflaLst)))

endTime = datetime.datetime.now()
print (endTime)
diffTime = endTime - startTime
print("Time for Completion: ", diffTime)
