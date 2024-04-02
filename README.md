# Email-request-from-purchase-data
This peace of code generates a pretty looking email request from what procurement guys usually get from SAP or 1S when digging up purchase requests.

>[!NOTE]
>To use your Gmail account for the script, first set up an app password in your Google account settings. Go to the App passwords, and under Select app, choose Mail, and under Select device choose Other (custom name) and fill in a name (such as “Purchase request”). Then put it in the script at my_email and password.


This is an example of an Excel file generated from interprice software like SAP or 1S.
In general what is represented here is purchase requests for Supply chain department to perform.
![image](https://github.com/Andrudewt/Email-request-from-purchase-data/assets/137271592/51d6e75a-3aed-457a-9c46-5fe36e13e4e4)
etc...
___

Next step is data sorting. The key moment is to sum up identical items and remove some spare details. 
The output table is going to look like this:

![image](https://github.com/Andrudewt/Email-request-from-purchase-request-data/assets/137271592/f6abc6dd-7756-4669-97e8-0df427952c24)

It's an HTML table, so it will support auto text rearrange.

As additional data there's an address book file with manufacturer/merchant emails.
![image](https://github.com/Andrudewt/Email-request-from-purchase-request-data/assets/137271592/ef7c29f2-00b5-450f-83f8-6a5fcdf0324e)

___

### As the result, the relevant recipients will get the purchase requests.

- ![Screenshot_1](https://github.com/Andrudewt/Email-request-from-purchase-request-data/assets/137271592/a89f1588-69c5-4fc4-91b1-bcbf06912655)

- ![Screenshot_2](https://github.com/Andrudewt/Email-request-from-purchase-request-data/assets/137271592/0847462c-07e9-4541-8be1-6ca1be217e93)

- ![Screenshot_3](https://github.com/Andrudewt/Email-request-from-purchase-request-data/assets/137271592/d0437234-52ad-4e35-9c70-56b9479a6d78)


Also in the input Excel file will be added two new sheets. One called "Not assigned", this for items with undefined manufacturer or None email in the address book. 
The second "Performed requests" will contain the number list of purchase requests with "email sent" status.
