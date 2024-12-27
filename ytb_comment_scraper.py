import pandas as pd
import requests
import re

API_KEY = "YOUR_YOUTUBE_DATA_API_KEY"  # Replace with your API key
EXCEL_FILE_EXTENSION = ".xlsx"

def get_video_id(video_url):
    """
    Extract the video ID from a YouTube URL.
    """
    pattern = r"(?:v=|\/)([0-9A-Za-z_-]{11}).*"
    match = re.search(pattern, video_url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid YouTube video URL.")

def get_video_title(video_id):
    """
    Fetch the video title using the YouTube Data API.
    """
    url = f"https://www.googleapis.com/youtube/v3/videos?id={video_id}&key={API_KEY}&part=snippet"
    response = requests.get(url)
    if response.status_code != 200:
        raise RuntimeError("Failed to fetch video title.")
    data = response.json()
    if "items" not in data or not data["items"]:
        raise RuntimeError("Video not found.")
    return data["items"][0]["snippet"]["title"]

def fetch_comments(video_id, max_results=100):
    """
    Fetch comments using the YouTube Data API with proper pagination.
    """
    comments = []
    url = f"https://www.googleapis.com/youtube/v3/commentThreads?key={API_KEY}&textFormat=plainText&part=snippet&videoId={video_id}&maxResults={max_results}"
    
    while url:
        response = requests.get(url)
        if response.status_code != 200:
            raise RuntimeError(f"Failed to fetch comments. Status code: {response.status_code}")
        
        data = response.json()
        
        # Process the current page of comments
        for item in data.get("items", []):
            comment = item["snippet"]["topLevelComment"]["snippet"]["textDisplay"]
            comments.append(comment)
        
        # Check for nextPageToken to continue fetching
        next_page_token = data.get("nextPageToken")
        if next_page_token:
            url = f"https://www.googleapis.com/youtube/v3/commentThreads?key={API_KEY}&textFormat=plainText&part=snippet&videoId={video_id}&maxResults={max_results}&pageToken={next_page_token}"
        else:
            # Exit the loop if no more pages are available
            break
    
    return comments


def save_comments_to_excel(comments, file_name):
    """
    Save comments to an Excel file.
    """
    if not comments:
        print("No comments to save.")
        return
    df = pd.DataFrame(comments, columns=["Comments"])
    df.to_excel(file_name, index=False)
    print(f"Comments saved to '{file_name}'")

def main():
    try:
        # Get video URL from user
        video_url = input("Enter the YouTube video URL: ").strip()
        video_id = get_video_id(video_url)
        print(f"Video ID: {video_id}")

        # Fetch video title
        video_title = get_video_title(video_id)
        print(f"Video Title: {video_title}")

        # Fetch comments
        print("Fetching comments...")
        comments = fetch_comments(video_id)
        print(f"Downloaded {len(comments)} comments.")

        # Save to Excel
        safe_video_title = re.sub(r'[<>:"/\\|?*]', '_', video_title)
        output_file_name = f"{safe_video_title}{EXCEL_FILE_EXTENSION}"
        save_comments_to_excel(comments, output_file_name)

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

