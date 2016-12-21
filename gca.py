from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from datetime import datetime, date, time
import argparse

parser = argparse.ArgumentParser(description='Google Calendar Admin.')
parser.add_argument('action', choices=['show','delete','move'],
                    help='what do you want to do with the calendar event(s) \
                    show, delete or move?')
parser.add_argument('-user', dest='inuser', 
                    help='the email address of the user that created the event(s)')
parser.add_argument('-event', dest='summary',
                    help='the title of the calendar event(s)')
parser.add_argument('-age', dest='inage', choices=['past','future','all'],
                    help='past, future or all events')
parser.add_argument('-dest', dest='destuser',
                    help='the email address of the user to move the event(s) to')
args = parser.parse_args()

# If modifying these scopes, update_event your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar \
https://www.google.com/calendar/feeds/default/owncalendars/full \
https://www.googleapis.com/auth/admin.directory.user.readonly'
CLIENT_SECRET_FILE = 'bis_client_secret.json'
APPLICATION_NAME = 'Google Calendar Admin'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'bis-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_domain_users(usrservice):
    """Gets all users from the domain to list specific calendar events

    Create a Google Admin SDK API service object and outputs a list of all
    users in the domain
    Creates a Google Calendar API service object and outputs a list of all 
    the calendar events for each user in the domain
    """
    usr_page_token = None
    while True:
        try:
            #retrives all domain users and places them into the results
            #object ordered by email address for further processing
            results = usrservice.users().list(customer='my_customer', 
            pageToken=usr_page_token, orderBy='email').execute()
            users = results.get('users', [])
            usr_page_token = results.get('nextPageToken')
            if not users:
                print('No users in the domain.')
                break
            else:
                #this loop processes each valid user and lists all calendar
                #events for that user to be checked for matching criteria
                for user in users:
                    prieml = user['primaryEmail']
                    #print(prieml, user['suspended'])
                    
                    if not user['suspended'] \
                    and user['includeInGlobalAddressList']:
                        get_cal_events(user, calservice)
            if not usr_page_token:
                break
        except ValueError:
            print('Oops!  Looks like there\'s an error.  Try again...')

def get_cal_events(user, calservice):
    """Gets all events per user from their primary calendar

    Loops through the events in the users primary calendar
    to return all events matching the specific criteria
    specified in the code
    """
    cal_page_token = None
    while True:
        try:
            #the next for loop retrives the calendar events
            #list to be checked for matching criteria
            prieml = user['primaryEmail']
            creator_to_del = 'sandra.jeanbaptiste@bis.gov.uk'
            event_to_del = 'Digital Directorate Team Meeting'
            events = calservice.events().list(calendarId=prieml,
                pageToken=cal_page_token).execute()
            for event in events['items']:
                if event['status'] != 'cancelled':
                    try:
                        #this is the criteri to be checked against
                        organiser = event['organizer']['email']
                        summary = event['summary']
                        if organiser == creator_to_del \
                        and summary == event_to_del:
                            try:
                                #checking for specific start date 
                                #in the event some events have different
                                #dateTime\date keywords
                                if event['start']['dateTime']:
                                    evdate = event['start']['dateTime']
                                    startDate = datetime.strptime(evdate[0:10],
                                    '%Y-%m-%d')
                                    today = datetime.today()
                                    if startDate > today:
                                        print('{0} ({1}) {2} {3}'.format(prieml,
                                            event['summary'],
                                            event['organizer']['email'],
                                            evdate[0:10]))
                            except KeyError:
                                #if the keyword is not dateTime 
                                #then fetch date keyword
                                evdate = event['start']['date']
                                startDate = datetime.strptime(evdate, '%Y-%m-%d')
                                today = datetime.today()
                                if startDate > today:
                                    print('{0} ({1}) {2} {3}'.format(prieml,
                                        event['summary'],
                                        event['organizer']['email'],
                                        evdate))
                    except KeyError:
                        continue
            cal_page_token = events.get('nextPageToken')
            if not cal_page_token:
                break
        except ValueError:
            print('Oops!  Thhe last event has an error.  Try again...')
            
def show_events(usrservice,calservice):
    """Gets all of the calendar events of the specified user
    
    Using the input args we've determined that this execution is to retrieve
    calendar events and we will no find those events that meet the criteria
    and return them to the display 
    """
    print(args.action, args.inuser, 'celendar events')
    
def delete_events(usrservice,calservice):
    """Delets all of the calendar events of the specified user
    
    Using the input args we've determined that this execution is to delete
    calendar events and we will no find those events that meet the criteria,
    delete them, and report to the display 
    """
    print(args.action, args.inuser, 'celendar events')


def move_events(usrservice,calservice):
    """Move all of the calendar events of the specified user to a new user
    
    Using the input args we've determined that this execution is to move
    calendar events from one user to another and report to the display 
    """
    print(args.action, args.inuser, 'celendar events to', args.destuser)

def main():
    """Gets all users from the domain to list specific calendar events

    Create a Google Admin SDK API service object and outputs a list of all users
    in the domain
    Creates a Google Calendar API service object and outputs a list of all the
    calendar events for each user in the domain
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    usrservice = discovery.build('admin', 'directory_v1', http=http)
    calservice = discovery.build('calendar', 'v3', http=http)

    if not args.inuser:
        parser.parse_args(['-h'])
    else:
        source_user = args.inuser.strip()
        if '@bis.gov.uk' in source_user:
            print(source_user)
        else:
            parser.parse_args(['-h'])
        
    if args.action == 'show':
        show_events(usrservice,calservice)
    if args.action == 'delete':
        delete_events(usrservice,calservice)        
    if args.action == 'move' \
    and args.destuser:
        move_events(usrservice,calservice)
    else:
        parser.parse_args(['-h'])
    
    #print(args.action)


if __name__ == '__main__':
    main()
