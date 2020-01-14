# GmailCalendar

Google calendar event generation program based on gmail message

## Install
it runs in python 3.6

pip install google-api-python-client
pip install google-auth-oauthlib==0.4.1

## How to use it
Google support the automatic calendar event generation mode already. However this function has limitations until now. 
In addition, 

To solve this problem, the program checks calendar event from gmail and generate calendar event.

The patterns to recognize google event from mail are based on Korean mail style.  
If you want to change this pattern, please check parse_info_from_gmail function.

```
주제; 
시간; 2020년 1월 1일 ~ 2020년 1월 8일
장소;
설명;

``` 
