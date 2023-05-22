import sys
import gspread
import googleapiclient.discovery
from gspread import Client, Spreadsheet
from gspread.exceptions import APIError
from googleapiclient.errors import HttpError
from google_sheets_editor_bot_helper import constants as const
from oauth2client.service_account import ServiceAccountCredentials


ACCESS_FRAME: [str] = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
                       'https://www.googleapis.com/auth/drive']  # Telling the bot what to access.


USER_JSON_KEY_FILE_NAME: str = input('\nPlease write the name of your json_key_file (with extension) & Press Enter: ')


try:
    CREDENTIALS_VALIDATION: ServiceAccountCredentials = ServiceAccountCredentials.from_json_keyfile_name(
        USER_JSON_KEY_FILE_NAME, ACCESS_FRAME)

except FileNotFoundError and OSError:
    if FileNotFoundError:
        print('\nNo such file or directory. Please provide a valid file name.')

    elif OSError:
        print('')
        raise
    input('\nPress Enter to Exit.')
    sys.exit(1)


AUTH_CLIENT: Client = gspread.authorize(CREDENTIALS_VALIDATION)  # Authorising our credentials.


USER_EMAIL_ADDRESS: str = input('\nPlease type your gmail address to get access to the spreadsheet & then '
                                'Press Enter: ')  # Giving the user's email, access to the new spreadsheet.


def writing_to_google_sheets(video_id: str, user_api_key: str) -> None:

    try:
        file_contents = read_from_utube_comments_section(video_id, user_api_key)
        file_columns: [str] = file_contents[0]
        file_rows: [dict[str, str]] = file_contents[1]

        print('\nCreating our Workbook...')
        viewership_comments_workbook: Spreadsheet = AUTH_CLIENT.create(f'U-tube Viewers\' Comments for Video-ID; '
                                                                       f'{video_id}.csv')
        viewership_comments_workbook.share(USER_EMAIL_ADDRESS, perm_type='user', role='writer')
        print('\n\tWorkbook created successfully.')

        print('\nWriting our Column Headers...')
        viewership_comments_workbook_sheet1 = AUTH_CLIENT.open(
            f'U-tube Viewers\' Comments for Video-ID; {video_id}.csv').sheet1
        viewership_comments_workbook_sheet1.insert_row(file_columns)
        print('\n\tColumn Headers written successfully.')

        try:
            print('\nWriting our row data...')
            [viewership_comments_workbook_sheet1.insert_row(list(file_row.values()), idx + 2) for idx, file_row in
             enumerate(file_rows)]
            print('\n\tGoogle Sheets file writing completed Successfully!')
        except APIError as api_err:
            print('\n', api_err)

    except TypeError as tperr:
        if TypeError:
            print(str(tperr))
        else:
            raise


def read_from_utube_comments_section(video_id: str, user_api_key: str) -> tuple[[str], [dict[str, str]]]:

    print('\nGetting the comments from YouTube...')

    comments_fetcher = googleapiclient.discovery.build(const.API_SERVICE_NAME, const.API_VERSION,
                                                       developerKey=user_api_key)

    try:
        frisbee = comments_fetcher.commentThreads().list(
            part="snippet",
            videoId=video_id,
            maxResults=30
        )

        PRODUCT = frisbee.execute()
        RESPONSE = PRODUCT

        ALL_COMMENTS: [] = []

        HEADINGS: [str] = ["AUTHOR", "COMMENT", "DATE", "TIME", "LIKES"]

        for key, _ in RESPONSE.items():

            if key == 'items':
                COMMENTS_COLLECTION = RESPONSE[
                    key]  # We get the value here. The "value" being the comment response i.e.
                # a dict containing all the comments and associated data of the video in question.

                for comment_bundle in COMMENTS_COLLECTION:  # Iterating to retrieve each comment & associated metrics.
                    author = comment_bundle['snippet']['topLevelComment']['snippet']['authorDisplayName']
                    comment = comment_bundle['snippet']['topLevelComment']['snippet']['textDisplay']
                    comment_date_time = comment_bundle['snippet']['topLevelComment']['snippet']['publishedAt']
                    formatted_comment_date = f'{comment_date_time.split("T")[0]}'
                    formatted_comment_time = f'{comment_date_time.split("T")[1].replace("Z", "")}'
                    comment_likes = comment_bundle['snippet']['topLevelComment']['snippet']['likeCount']

                    ALL_COMMENTS.append({"AUTHOR": author, "COMMENT": comment, "DATE": formatted_comment_date,
                                         "TIME": formatted_comment_time, "LIKES": comment_likes})

        print('\n\tYouTube comments fetching completed successfully.\n')

        return HEADINGS, ALL_COMMENTS

    except HttpError and TypeError:
        if 'parameter could not be found' in str(HttpError):
            print('\n\t[ERR-404]: The video could not be found. Please provide a valid link.')
        elif HttpError:
            print(str(HttpError))
        elif TypeError:
            print(str(TypeError))
