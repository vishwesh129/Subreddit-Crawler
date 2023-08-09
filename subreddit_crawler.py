import requests
from bs4 import BeautifulSoup
import pandas as pd

def get_top_comments(subreddit_name):
    after_id = ""
    data = []
    try:
        # if after_id:
        #   url = f"https://old.reddit.com//r/{subreddit_name}/top/?sort=top&t=all&after={after_id}"
        # else:
        #   url = f"https://old.reddit.com//r/{subreddit_name}/top/?sort=top&t=all"
        
        url = f"https://old.reddit.com//r/{subreddit_name}/top/?sort=top&t=all"
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            posts = soup.find_all("div", {"class": "thing"})
            if posts:
                for post in posts:
                    # Skipping promoted posts
                    if "promoted" not in post.get("class"):
                        
                        # parsing title of the post
                        entry = post.find("div", {"class": "entry"})
                        title = entry.find("a", {"class": "title"}).text
                        print("title ---->>>", title + "\n\n")
                        
                        # parsing link, upvotes count and comments count
                        upvotes_count = post.get("data-score")
                        comments_count =  post.get("data-comments-count")
                        outlink = post.get("data-url")
                        after_id = post.get("data-fullname")

                        # For getting the description and top comment of each post
                        post_url = "https://old.reddit.com" + post.get("data-permalink") + "?sort=top&t=all"
                        post_response = requests.get(post_url, headers=headers)
                        post_soup = BeautifulSoup(post_response.text, 'html.parser')

                        # Parse description
                        description = ""
                        expando = post_soup.find("div", {"class":"expando"})
                        if expando and expando.find_all("div", {"class": "usertext-body"}):
                            description = expando.find_all("div", {"class": "usertext-body"})[0].text
                        
                        # Parse top comment of the post
                        top_comment_text = ""
                        comments = post_soup.find("div", {"class": "commentarea"})
                        if comments:
                            top_comment = comments.find_all("div", {"class": "usertext-body"})[0]
                            top_comment_text = top_comment.find("p").text
                        data.append([title, description, upvotes_count, comments_count, outlink, top_comment_text])
            else:
              print("No subreddit exists - ", subreddit_name)
        else:
            print("Error while browsing the subreddit - ", subreddit_name)
            return []
    except Exception as e:
        print(f"Error: {e}")
        return []
    return data

def create_excel(data, subreddit_name):
    df = pd.DataFrame(data, columns=["Title", "Description", "Upvotes", "Number of Comments", "Outlink", "Top Comment"])
    excel_file = f"{subreddit_name}_top_posts.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')
    print(f"Excel file '{excel_file}' created successfully!")

if __name__ == "__main__":
    # Read subreddit name from user
    subreddit_name = input("Enter subreddit: ").strip()
    subreddit_data = get_top_comments(subreddit_name)
    
    # Create excel sheet when there is data present
    if subreddit_data:
        create_excel(subreddit_data, subreddit_name)

