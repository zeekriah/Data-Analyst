#Importing libaries
import warnings 
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np

# Loading data from Excel
Order = pd.read_excel("Company X - Order Report.xlsx")
pincode = pd.read_excel("Company X - Pincode Zones.xlsx")
sku = pd.read_excel("Company X - SKU Master.xlsx")
invoice = pd.read_excel("Courier Company - Invoice.xlsx")

#merging data by using merge function
oi_merge = pd.merge(Order,invoice,on="Order_ID",how="right")
ois_merge = pd.merge(oi_merge,sku,on="SKU",how="left")
final_merge = pd.merge(ois_merge,pincode,on="Customer Pincode",how="inner")

#column renaming
final_merge.rename(columns={"Warehouse_Pincode_x":"Company_Warehouse_Pincode",
                     "Customer Pincode":"Company_Customer_Pincode","Zone_x":"Company_Zone",
                     "Warehouse Pincode":"X_Warehouse_Pincode","Zone_y":"X_Zone","Weight (g)":"Weight(g)"},inplace=True)
final_merge.rename(columns={"Company_Zone":"Delivery_Zone_charged_by_Courier_Company","X_Zone":"Delivery_Zone_as_per_X"},inplace=True)
final_merge.rename(columns={"Charged Weight":"Total_weight_as_per_Courier_Company_(KG)"},inplace=True)

#multiplication of order qty*weight(g) to find out total weight as per x(g)
final_merge["Total_weight_as_per_X(g)"]=final_merge["Order Qty"]*final_merge["Weight(g)"]

#dividing the total weight as per X(g)/1000,So we get total weight as per x(kg)
final_merge["Total_weight_as_per_X(Kg)"]=final_merge["Total_weight_as_per_X(g)"]/1000

#to find out weight_slab_as_per_X_company,we are using if statement/condition
condition=[(final_merge["Total_weight_as_per_X(Kg)"]<=0.500),
          (final_merge["Total_weight_as_per_X(Kg)"]>0.500)&(final_merge["Total_weight_as_per_X(Kg)"]<=1),
          (final_merge["Total_weight_as_per_X(Kg)"]>1)&(final_merge["Total_weight_as_per_X(Kg)"]<=1.5),
          (final_merge["Total_weight_as_per_X(Kg)"]>1.5)&(final_merge["Total_weight_as_per_X(Kg)"]<=2),
          (final_merge["Total_weight_as_per_X(Kg)"]>2)&(final_merge["Total_weight_as_per_X(Kg)"]<=2.5),
          (final_merge["Total_weight_as_per_X(Kg)"]>2.5)&(final_merge["Total_weight_as_per_X(Kg)"]<=3),
          (final_merge["Total_weight_as_per_X(Kg)"]>3)&(final_merge["Total_weight_as_per_X(Kg)"]<=3.5),
          (final_merge["Total_weight_as_per_X(Kg)"]>3.5)&(final_merge["Total_weight_as_per_X(Kg)"]<=4)]
value=[0.5,1,1.5,2,2.5,3,3.5,4]
final_merge["Weight_slab_as_per_X_(KG)"]=np.select(condition,value)
print(final_merge)

#To find out Weight_slab_charged_by_Courier_Company_(KG) again use if condtn/statement
conditions=[(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=0.500),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>0.500)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=1),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>1)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=1.5),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>1.5)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=2),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>2)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=2.5),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>2.5)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=3),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>3)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=3.5),
          (final_merge["Total_weight_as_per_Courier_Company_(KG)"]>3.5)&(final_merge["Total_weight_as_per_Courier_Company_(KG)"]<=4)]
value=[0.5,1,1.5,2,2.5,3,3.5,4]
final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]=np.select(conditions,value)


#If condition- To find out Expected_Charge_as_per_X_(Rs.)
conditionss=[(final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="b")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="d")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_as_per_X"]=="e")&
             (final_merge["Weight_slab_as_per_X_(KG)"]==1.5)]
             


valuess=[(final_merge["Billing Amount (Rs.)"]+33),(final_merge["Billing Amount (Rs.)"]+61.3),(final_merge["Billing Amount (Rs.)"]+89.6),
         (final_merge["Billing Amount (Rs.)"]+45.4),(final_merge["Billing Amount (Rs.)"]+90.2),(final_merge["Billing Amount (Rs.)"]+135),
        (final_merge["Billing Amount (Rs.)"]+56.6),(final_merge["Billing Amount (Rs.)"]+112.1),(final_merge["Billing Amount (Rs.)"]+167.6),
        (final_merge["Billing Amount (Rs.)"]+45.4),(final_merge["Billing Amount (Rs.)"]+110.1),(final_merge["Billing Amount (Rs.)"]+166.7),
         (final_merge["Billing Amount (Rs.)"]+86.7),(final_merge["Billing Amount (Rs.)"]+176.3),(final_merge["Billing Amount (Rs.)"]+265.9),
         (final_merge["Billing Amount (Rs.)"]+107.3),(final_merge["Billing Amount (Rs.)"]+218.3),(final_merge["Billing Amount (Rs.)"]+329.3)]

final_merge["Weight_slab_as_per_X_(KG)"].value_counts()
final_merge["Delivery_Zone_as_per_X"].value_counts()
final_merge["Type of Shipment"].value_counts()
final_merge["Expected_Charge_as_per_X_(Rs.)"]=np.select(conditionss,valuess)

#If condition- To find out the Charges_Billed_by_Courier_Company_(Rs.)
con=[(final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="b")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5), 
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2), 
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="d")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==0.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.0),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==1.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==2.5),
             (final_merge["Type of Shipment"]=="Forward and RTO charges")&(final_merge["Delivery_Zone_charged_by_Courier_Company"]=="e")&
             (final_merge["Weight_slab_charged_by_Courier_Company_(KG)"]==3)]

val=[(final_merge["Billing Amount (Rs.)"]+33),(final_merge["Billing Amount (Rs.)"]+61.3),(final_merge["Billing Amount (Rs.)"]+89.6),(final_merge["Billing Amount (Rs.)"]+117.9),(final_merge["Billing Amount (Rs.)"]+146.2),(final_merge["Billing Amount (Rs.)"]+174.5),(final_merge["Billing Amount (Rs.)"]+45.4),(final_merge["Billing Amount (Rs.)"]+90.2),(final_merge["Billing Amount (Rs.)"]+135),(final_merge["Billing Amount (Rs.)"]+179.8),(final_merge["Billing Amount (Rs.)"]+224.6),(final_merge["Billing Amount (Rs.)"]+269.4),(final_merge["Billing Amount (Rs.)"]+56.6),(final_merge["Billing Amount (Rs.)"]+112.1),(final_merge["Billing Amount (Rs.)"]+167.6),(final_merge["Billing Amount (Rs.)"]+223.1),(final_merge["Billing Amount (Rs.)"]+278.6),(final_merge["Billing Amount (Rs.)"]+334.1),(final_merge["Billing Amount (Rs.)"]+45.4),(final_merge["Billing Amount (Rs.)"]+110.1),(final_merge["Billing Amount (Rs.)"]+166.7),(final_merge["Billing Amount (Rs.)"]+223.3),(final_merge["Billing Amount (Rs.)"]+279.9),(final_merge["Billing Amount (Rs.)"]+336.5),(final_merge["Billing Amount (Rs.)"]+86.7),(final_merge["Billing Amount (Rs.)"]+176.3),(final_merge["Billing Amount (Rs.)"]+265.9),(final_merge["Billing Amount (Rs.)"]+355.5),(final_merge["Billing Amount (Rs.)"]+445.1),(final_merge["Billing Amount (Rs.)"]+534.7),(final_merge["Billing Amount (Rs.)"]+107.3),(final_merge["Billing Amount (Rs.)"]+218.3),(final_merge["Billing Amount (Rs.)"]+329.3),(final_merge["Billing Amount (Rs.)"]+440.3),(final_merge["Billing Amount (Rs.)"]+551.3),(final_merge["Billing Amount (Rs.)"]+662.3)]

final_merge["Weight_slab_charged_by_Courier_Company_(KG)"].value_counts()
final_merge["Delivery_Zone_charged_by_Courier_Company"].value_counts()
final_merge["Charges_Billed_by_Courier_Company_(Rs.)"]=np.select(con,val)

#finding out Difference_Between_Expected_Charges_and_Billed_Charges_(Rs.) & calculations
final_merge["Difference_Between_Expected_Charges_and_Billed_Charges_(Rs.)"]= final_merge["Expected_Charge_as_per_X_(Rs.)"]- final_merge["Charges_Billed_by_Courier_Company_(Rs.)"]

#assing object
sheet=final_merge

#deleting unwanted columns
delete=["SKU","Order Qty","Company_Warehouse_Pincode","Company_Customer_Pincode","Type of Shipment",
      "Billing Amount (Rs.)","Weight(g)","X_Warehouse_Pincode","Total_weight_as_per_X(g)"]

#using for loop to delete
for i in delete:
    sheet.drop([i],inplace=True,axis=1)
    
#Arrange columns as per submission guidelines
FINAL_SHEET=sheet[["Order_ID","AWB Code","Total_weight_as_per_X(Kg)","Weight_slab_as_per_X_(KG)",
                     "Total_weight_as_per_Courier_Company_(KG)","Weight_slab_charged_by_Courier_Company_(KG)",
                     "Delivery_Zone_as_per_X","Delivery_Zone_charged_by_Courier_Company","Expected_Charge_as_per_X_(Rs.)",
                    "Charges_Billed_by_Courier_Company_(Rs.)","Difference_Between_Expected_Charges_and_Billed_Charges_(Rs.)"]]

#Assign Final submission Sheet to Final Object
FINAL=FINAL_SHEET

# To find out the "Total orders where X has been correctly charged"
cK=[(FINAL["Charges_Billed_by_Courier_Company_(Rs.)"])==(FINAL["Expected_Charge_as_per_X_(Rs.)"])]
VO=[FINAL["Expected_Charge_as_per_X_(Rs.)"]]
FINAL["Total orders where X has been correctly charged"]=np.select(cK,VO)

# To find out the "Total Orders where X has been overchargeds"
K=[(FINAL["Charges_Billed_by_Courier_Company_(Rs.)"])>(FINAL["Expected_Charge_as_per_X_(Rs.)"])]
O=[FINAL["Difference_Between_Expected_Charges_and_Billed_Charges_(Rs.)"]]
FINAL["Total Orders where X has been overchargeds"]=np.select(K,O)

# To find out the Total Orders where X has been undercharged"
Kw=[(FINAL["Charges_Billed_by_Courier_Company_(Rs.)"])<(FINAL["Expected_Charge_as_per_X_(Rs.)"])]
Ow=[FINAL["Difference_Between_Expected_Charges_and_Billed_Charges_(Rs.)"]]
FINAL["Total Orders where X has been undercharged"]=np.select(Kw,Ow)

#Converting python file into Excel file & excel stored in below location with output data in it.
FINAL.to_excel("C:\\Users\\hussainzeekria\\Output 1-2.xlsx")
