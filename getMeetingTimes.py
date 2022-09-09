#!python3
# A script to process and format available meeting times in Outlook
# Version 1.0 8/22/22

import requests
import json
from datetime import datetime, timedelta

# input data is retrieved from Power Automate "Get meeting times" cloud flow via HTTP request
# input = ["2022-08-23T08:00:00.0000000", "2022-08-25T11:00:00.0000000", "2022-08-26T11:30:00.0000000",
#        "2022-08-29T08:00:00.0000000", "2022-08-29T12:00:00.0000000", "2022-08-31T11:30:00.0000000",
#        "2022-08-23T12:00:00.0000000", "2022-08-25T15:00:00.0000000", "2022-08-26T15:30:00.0000000",
#        "2022-08-29T12:00:00.0000000", "2022-08-29T16:00:00.0000000", "2022-08-31T15:30:00.0000000"]

# overall result using above input times:
# 8/23:  8 - 12 CST
# 8/25: 11 - 3 CST
# 8/26: 12 - 3 CST
# 8/29:  8 - 4 CST
# 8/31: 12 - 3 CST

# Settings to query power automate cloud flow HTTP listener URL
cloud_url = 'https://prod-34.westus.logic.azure.com:443/workflows/0b0f0bd35618401f86477a7a830f1dac/triggers/manual/paths/invoke?api-version=2016-06-01&sp=/triggers/manual/run&sv=1.0&sig=gtV0akNrZOZic_icEBuOWGSzvtcWn8CdMH_IgShHyhY'
headers = {'Content-Type': 'application/json'}
# http request body below must be wrapped in ''' to avoid error for invalid characters like the comma
# modify the email and duration as needed
body = '''{'email':'joel_ginsberg@trendmicro.com','duration':240}'''

# make HTTP request and store response in r
r = requests.post(cloud_url, headers=headers, data=body)
# parse json data from HTTP response content and assign it to input array
input = json.loads(r.content)
print(input)

# sort times chronologically
input.sort()

# combine consecutive meeting blocks into one to make output more intuitive
# the times list contains both start and end times in chronological order, evens are start times and odds are end times
# for example if there are three 1-hour long meeting blocks:
# 8am - 9am, 9am - 10am, and 10am - 11am, we want to combine them into one block: 8am - 11am
# we have to call the function recursively until all consecutive meeting blocks have been combined
# all consecutive meeting blocks have been combined if the function runs and does not remove any elements from the list
def combineTimes(times):
    # get initial length of times list to compare to length after running this function
    # this is how we check if we can stop calling the function recursively
    beforeLength = len(times)

    for i, e in enumerate(times):
        # avoid error on last element
        if i < (len(times) - 1):
            # check if current datetime is the same as the next one
            # print input[i], input[i+1]
            if times[i] == times[i + 1]:
                # after the first pop the next element [i+1] becomes the current one [i]
                # we want to remove both duplicates
                times.pop(i)
                times.pop(i)

    # check if all consecutive meeting times have been combined
    if len(times) == beforeLength:  # if no elements were removed then all consecutive times have been combined
        return times
    else:
        # return is necessary to store result of recursive function call, otherwise first function call returns None
        return combineTimes(times)  # call function again until all consecutive times have been combined

input = combineTimes(input)

# print(input)


# Change :30 times to :00. To avoid conflict with other meetings, add 30 min to start times and subtract 30 from end times
# start times should always be even numbers (starting with 0) and end times should be odd numbers in the list
def shrinkby30(times):
    # add 30 min to start times
    for i, e in enumerate(times[0::2]):  # check every two elements in list starting with first element
        if ":30" in e:
            dt = datetime.strptime(e, '%Y-%m-%dT%H:%M:%S.0000000')  # convert string to datetime
            dt = dt + timedelta(minutes=30)  # add 30 min
            # multiply i * 2 because enumerate only got every other element from times so it's half as long
            times[i * 2] = dt.strftime('%Y-%m-%dT%H:%M:%S.0000000')  # format dt as string again

    # subtract 30 min from end times
    for i, e in enumerate(times[1::2]):  # check every two elements in list starting with second element
        if ":30" in e:
            dt = datetime.strptime(e, '%Y-%m-%dT%H:%M:%S.0000000')
            dt = dt - timedelta(minutes=30)  # subtract 30 min
            # multiply i * 2 + 1 bc enumerate got every other element from times starting with position 1 not 0
            times[i * 2 + 1] = dt.strftime('%Y-%m-%dT%H:%M:%S.0000000')

    return times


input = shrinkby30(input)

# Combine every two items in array with ' - ' so we can show time span like 8 - 4 pm
# https://stackoverflow.com/questions/24443995/list-comprehension-joining-every-two-elements-together-in-a-list
output = []
for x, y in zip(input[0::2], input[1::2]):
    # convert strings x and y to datetime objects a and b so they can be formatted for human readable output
    a = datetime.strptime(x, '%Y-%m-%dT%H:%M:%S.0000000')
    b = datetime.strptime(y, '%Y-%m-%dT%H:%M:%S.0000000')
    # format list with human readable day of the week, month/day, and times plus time zone
    # dt.strftime("%m/%d %I:%M%p %Z")
    # I hard-coded CST instead of printing the timezone with %Z because the datetime objects returned from the cloud flow do not have a timezone value.
    #output.append(a.strftime("%m/%d:").lstrip('0').replace('/0', '/') + a.strftime(" %I").replace(' 0', ' ').rjust(3, ' ') + ' - ' + b.strftime("%I CST").lstrip('0'))
    output.append(a.strftime("%m/%d:").lstrip('0').replace('/0', '/') + a.strftime(" %I %p").replace(' 0', ' ').lower() + ' - ' + b.strftime("%I %p").lstrip('0').lower() + " CST")

print('\n'.join(output))
