import win32com.client


def get_appointments(begin, end):
    """
    Queries Outlook for appointment items between the specified begin and end dates.
    Args:
        begin (datetime.datetime): The start time of the interval to query, as a `datetime.datetime` object.
        end (datetime.datetime): The end time of the interval to query, as a `datetime.datetime` object.
    Returns:
        str: A string containing a formatted list of appointment items.
    """

    # get a handle on Outlook
    outlook = win32com.client.Dispatch(
        'Outlook.Application').GetNamespace('MAPI')

    # 9 = list of all the meetings, 6 = emails
    calendar = outlook.getDefaultFolder(9).Items

    # events that repeat
    calendar.IncludeRecurrences = True

    # sort
    calendar.Sort("[Start]")

    # setup a constraint
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime(
        '%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)

    # show what we found in Outlook
    for appointment in calendar:
        print("Day of month: ", appointment.StartInStartTimeZone.day)
        print("Subject: ", appointment.Subject)
        # print("Is recurring: ", appointment.IsRecurring)
        # print("Is conflicted: ", appointment.IsConflict)
        # print("Is reminder set: ", appointment.ReminderSet)
        print("Start: ", appointment.Start)
        print("---------")

    return calendar

def groom_appointments(calendar):
    """
    Takes a list of Outlook appointment items and processes them into a dictionary format suitable for display.
    Args:
        calendar (list): A list of Outlook appointment items.
    Returns:
        None
    """
    appointmentDictionary = {}

    for appointment in calendar:
        meetingDate = str(appointment.Start)
        subject = str(appointment.Subject)
        duration = str(appointment.duration)
        # date = parse(meetingDate).date()
        # time = parse(meetingDate).time()
        appointmentDictionary[subject] = {"Subject": [
            subject], "Time": [meetingDate], "Durations": [duration]}

    for subject in appointmentDictionary.keys():
        rowDict = {}
        rowDict["Subject"] = appointmentDictionary[subject]["Subject"] if appointmentDictionary[subject]["Subject"] else ""