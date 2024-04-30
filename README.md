# Overview

This is a Robocorp automation robot for extracting news articles from a website and saving them to an Excel file.

## Overview

This automation robot performs the following tasks:
- Opens a website (https://www.aljazeera.com/search/obama?sort=date)
- Inserts a search query retrieved from Control Room's cloud workspace
- Retrieves a list of articles based on the number of months specified in the cloud
- Creates an Excel file
- Inserts rows into the Excel file containing information about each article
- Downloads the image for each processed article

Input data: Two arguments from the cloud, named "SearchPhrase" and "NumberOfMonths"
Output (if needed): All the pictures from the processed articles, log.html, and the Excel file

Link for exercise overview: [RPA Challenge - Fresh news 2.0](https://thoughtfulautomation.notion.site/RPA-Challenge-Fresh-news-2-0-37e2db5f88cb48d5ab1c972973226eb4)

## Results

🚀 After running the bot, check out the `output` folder:
- An Excel file is generated, named "news_report.xlsx"
- Check `log.html` for any logged information
- Review the images downloaded from articles within the specified time constraints  

## Quick reminders

- This automation employs a pagination mechanism. While the articles meet the date requirement, the automation continues to click "Show more" to expand the list of articles.
- The news page (https://www.aljazeera.com/search/obama?sort=date) is based in Qatar, so all date manipulations consider their timezone.
- The news page does not include category filtering after the search phrase.
- Each image is downloaded using the 'alt' tag from HTML. Sometimes this attribute does not exist, so the file name is set as "Article" followed by the number of the processed item.
- If the search phrase does not retrieve any data, the automation ends successfully without generating output data in the `output` folder.
- The maximum number of retries to open the browser on the news URL is set to 3.

