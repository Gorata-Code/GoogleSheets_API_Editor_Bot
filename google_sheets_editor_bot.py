import sys
from google_sheets_editor_bot_helper.g_sheets_editor import writing_to_google_sheets


def script_summary() -> None:
    print('''
               ***----------------------------------------------------------------------------------------***
         \t***------------------------ DUMELANG means GREETINGS! ~ G-CODE -----------------------***
                     \t***------------------------------------------------------------------------***\n
              
        \t"GOOGLE SHEETS-API-EDITOR-BOT" Version 1.0.0\n
        
        Run this script to scrape online data and write it to a Google Sheets spreadsheet document.
        Simply provide a YouTube video link, your API key and the script will fetch the video comments for you,
        and then it will save them to a Google Sheet.
        
        Peace!
        
    ''')


def sheets_bot(video_id: str, user_api_key: str) -> None:
    try:
        writing_to_google_sheets(video_id, user_api_key)

    except Exception as exp:

        if 'INTERNET' in str(exp):

            print(''''

                            Please make sure you are connected to the internet and Try again.

                            Cheers!

                            ''')

            input('\nPress Enter To Exit.\n')

        elif 'Timed out receiving message from renderer' or 'cannot determine loading status' in str(exp):
            print('\nIt appears you\'re experiencing internet connectivity issues.')
            print(str(exp))

        elif 'ERR_NAME_NOT_RESOLVED' or 'ERR_CONNECTION_CLOSED' or 'unexpected command response' in str(exp):
            print('\nYour internet connection may have been interrupted.')
            print('Please make sure you\'re still connected to the internet before you try again.')

        input('\nPress Enter to Exit & Try Again.')
        sys.exit(1)

    input('\nPress Enter to Exit.')
    sys.exit(0)


def main() -> None:
    script_summary()

    video_id: str = input('\nPlease paste the youtube video link here: ').strip().split("/")[-1].replace("watch?v=", "")
    user_api_key: str = input('\nPlease paste your api_key_here: ').strip()

    if len(video_id) >= 1 and len(user_api_key) >= 1:
        sheets_bot(video_id, user_api_key)
    elif len(video_id) < 1 or len(user_api_key) < 1:
        print('\nPlease provide all the required information.')
        input('\nPress Enter to Exit: ')
        sys.exit(1)


if __name__ == '__main__':
    main()
