import tweepy
import pandas as pd
import time

# Twitter API credentials
consumer_key = 'key_from_twitter'
consumer_secret = 'secret_from_twitter'
access_token = 'token_from_twitter'
access_token_secret = 'token_secret_from_twitter'

# Authenticate to Twitter
auth = tweepy.OAuth1UserHandler(consumer_key, consumer_secret, access_token, access_token_secret)
api = tweepy.API(auth)

# Read data from Excel sheet
excel_file = 'tweets.xlsx'
df = pd.read_excel(excel_file)

# Function to post tweet
def post_tweet(tweet):
    try:
        api.update_status(tweet)
        print("Tweet posted successfully!")
    except tweepy.TweepError as e:
        print(f"Error: {e}")

# Main loop to post tweets periodically
for index, row in df.iterrows():
    tweet_content = row['Tweet']
    post_tweet(tweet_content)
    time.sleep(3600)  # Sleep for 1 hour (adjust as needed)
