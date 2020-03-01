# Auto Reset Password for SC2

read sc2 temp password from outlook and auto reset it by RPA

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them
1. Python version 3.6+
```
https://www.python.org/downloads/release/python-370/
```
2. Visual Studio Code
```
https://code.visualstudio.com/download
```
3. Outlook Folder structure (SC2 Folder inside inbox)
```
Inbox (Parent)
->SC2 (Subfolder)
```
### Installing

A step by step series of examples that tell you how to get a development env running

1. Clone this repository to your local

```
git clone [THIS REPOSITORY URL]
```

2. run following pip commands

```
pip install rpa
```
```
pip install pandas
```
```
pip install pyOutlook
```

Note : If found error like "No module found named .??!" Please use python -m pip install xxx instead of pip Example
```
pip install pyOutlook
```

## Running the tests

After finished prerequisite installation, Kindly do following steps
1. Get access token by follow link 
```
https://oauthplay.azurewebsites.net
```
Then, copy and replace code [YOUR_ACCOUNT_ACCESS_TOKEN] with your access token
This token needed to renew every 1 hour.

2.Run code once, 
```
python test2.py
```
and Get copy your inbox folder "id" , replace code [YOUR INBOX FOLDER ID] with your inbox folder ID
```
InboxFolder = account_one.get_folder_by_id([YOUR INBOX FOLDER ID])
```
3.Run code once, 
```
python test2.py
```
and Get copy your SC2 folder "id" , replace code [YOUR SC2_FOLDER ID] with your SC2 folder ID
```
SC2Folder = account_one.get_folder_by_id('YOUR_SC2_FOLDER_ID')
```

4.Run it the last time to create temp excel file to store temp user/pass from your outlook SC2 folder
```
python test2.py
```

5.Run robot
```
python test2_keyin.py
```


## Authors

* **Thitipong Kanjanapa** - *Initial work* - [knotict](https://github.com/knotict)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Just for fun, productivity improvement
* Do less, result more
* etc
