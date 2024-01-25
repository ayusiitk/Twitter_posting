import tweepy
import pandas as pd
import time

# Twitter API credentials
consumer_key = 'your_consumer_key'
consumer_secret = 'your_consumer_secret'
access_token = 'your_access_token'
access_token_secret = 'your_access_token_secret'

# Authenticate to Twitter
auth = tweepy.OAuth1UserHandler(consumer_key, consumer_secret, access_token, access_token_secret)
api = tweepy.API(auth)

# Function to retrieve top trending topic
def get_top_trending_topic():
    try:
        trends = api.trends_place(1)  # Get trending topics for worldwide
        top_trend = trends[0]['trends'][0]['name']
        return top_trend
    except tweepy.TweepError as e:
        print(f"Error: {e}")
        return None

# Function to append tweet to Excel file
def append_tweet_to_excel(tweet):
    try:
        df = pd.read_excel('tweets.xlsx')
        df = df.append({'Tweet': tweet}, ignore_index=True)
        df.to_excel('tweets.xlsx', index=False)
        print("Tweet appended to Excel file successfully!")
    except Exception as e:
        print(f"Error: {e}")

# Main loop to fetch top trending topic and append tweet every hour
while True:
    top_trend = get_top_trending_topic()
    if top_trend:
        tweet_content = f"The top trending topic is: {top_trend} #trendingtopic"
        append_tweet_to_excel(tweet_content)
    time.sleep(3600)  # Sleep for 1 hour
