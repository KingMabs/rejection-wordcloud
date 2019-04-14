from __future__ import print_function
import math, pickle, os.path, base64, re, pathlib, validators
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import errors
from collections import Counter
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


# ========================= Gmail Functions =========================

def ListMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError as error:
    print('An error occurred: %s' % error)
    
def GetMessage(service, user_id, msg_id):
  """Get a Message with given ID.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A Message.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    return message
  except errors.HttpError as error:
    print('An error occurred: %s' % error)

# ========================= Utility Functions ========================= 

class FormatCounter(Counter):
    def __str__(self):
        return "\n".join('{} {}'.format(k, v) for k, v in sorted(self.items(), key=lambda item: (-item[1], item[0])))

def ProgressBar(msg, iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█'):
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    os.system('cls')
    print(msg)
    print("\r%s |%s| %s%% %s \n" % (prefix, bar, percent, suffix), end = '\r', flush=True)

# =============================== Constants =============================== 

USER_ID = 'me'
ID_STR, PAYLOAD_STR, PARTS_STR, MIMTYPE_STR, BODY_STR, DATA_STR = 'id', 'payload', 'parts', 'mimeType', 'body', 'data'

CRED_FILE, OUTPUT_FILE, OUTPUT_IMAGE_FILE  = 'credentials.json', 'output.txt', 'output.png'

LABELS = ['jobs-2018-rejections', 'jobs-2019-rejections']

IMG_WIDTH , IMG_HEIGHT = 800, 400

# =============================== Main =============================== 

def main():  
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CRED_FILE, SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # ------- Get rejection emails -------
    query_str = '{'
    for label in LABELS:
      query_str += f'label:{label} '
    query_str += '}'  

    rejection_emails = ListMessagesMatchingQuery(service, USER_ID, query_str)
    
    wordList = []

    prog_message = 'Reading rejection emails...'
    prog_prefix, prog_suffix = 'Progress:', 'Complete'
    completion_val, total_val = 100, len(rejection_emails)
    prog_length = int(completion_val / 2)
    count, iter_step = 0, int(math.ceil(total_val / completion_val))
    
    stopwords = set(STOPWORDS)

    # ------- Read rejection emails -------
    for i in range(total_val):
      email = rejection_emails[i]
      message = GetMessage(service, USER_ID, email[ID_STR])
      msg_str = ''
      partsList = message[PAYLOAD_STR][PARTS_STR] if PARTS_STR in message[PAYLOAD_STR] else [message[PAYLOAD_STR]]
      for part in partsList:
        if part[MIMTYPE_STR] == 'text/plain':
          part_data = part[BODY_STR][DATA_STR]
          part_str = str(base64.urlsafe_b64decode(part_data.encode('ASCII')))
          decoded_part_str = bytes(part_str, 'utf-8').decode('unicode_escape')
          msg_str += decoded_part_str
      if msg_str:
        count += 1
      msg_str = msg_str[2:]

      for word in msg_str.split():
        word = ''.join(word.strip().strip('[-()\"#/@;:<>{}`+=~|.!?,]').split())
        is_url_or_email = validators.url(word) or validators.email(word)
       
        if len(word) > 1 and not is_url_or_email:
          word = re.sub('[^ a-zA-Z0-9]', '', word.lower())
          if word and word not in stopwords and not word.isdigit() and not (word.startswith('http') or word.startswith('www') or word.endswith('.com')):
            wordList.append(word)
      if i % iter_step == 0:
        ProgressBar(prog_message, i, total_val, prefix = prog_prefix, suffix = prog_suffix, length = prog_length)
    
    if (total_val > 0): ProgressBar(prog_message, total_val, total_val, prefix = prog_prefix, suffix = prog_suffix, length = prog_length)    
    print(f'Read {count}/{total_val} rejection emails', end='\n\n')

    # ------- Generate files -------
    counter = FormatCounter(wordList)

    print('Generating word cloud image...')
    WordCloud(stopwords=stopwords,width=IMG_WIDTH,height=IMG_HEIGHT).generate_from_frequencies(counter).to_file(OUTPUT_IMAGE_FILE)
    pathlib.Path(OUTPUT_FILE).write_text(str(counter))
    print('Complete!')
            
if __name__ == '__main__':
    main()
