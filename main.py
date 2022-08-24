# A script to process and format available meeting times in Outlook
# Version 1.0 8/22/22

from datetime import datetime, timedelta

# input data is retrieved from Power Automate "Get meeting times" cloud flow via HTTP request
input = ["2022-08-23T08:00:00.0000000", "2022-08-25T11:00:00.0000000", "2022-08-26T11:30:00.0000000",
         "2022-08-29T08:00:00.0000000", "2022-08-29T12:00:00.0000000", "2022-08-31T11:30:00.0000000",
         "2022-08-23T12:00:00.0000000", "2022-08-25T15:00:00.0000000", "2022-08-26T15:30:00.0000000",
         "2022-08-29T12:00:00.0000000", "2022-08-29T16:00:00.0000000", "2022-08-31T15:30:00.0000000"]

# overall result using above input times:
# Tuesday 08/23 08:00AM - 12:00PM CST
# Thursday 08/25 11:00AM - 03:00PM CST
# Friday 08/26 12:00PM - 03:00PM CST
# Monday 08/29 08:00AM - 04:00PM CST
# Wednesday 08/31 12:00PM - 03:00PM CST


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
            #print input[i], input[i+1]
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
#print input


# Change :30 times to :00. To avoid conflict with other meetings, add 30 min to start times and subtract 30 from end times
# start times should always be even numbers (starting with 0) and end times should be odd numbers in the list
def shrinkby30(times):
    # add 30 min to start times
    for i, e in enumerate(times[0::2]):  # check every two elements in list starting with first element
        if ":30" in e:
            dt = datetime.strptime(e, '%Y-%m-%dT%H:%M:%S.0000000')   # convert string to datetime
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
    # dt.strftime("%A %m/%d %I:%M%p %Z")
    output.append(a.strftime("%a %m/%d %I:%M%p").replace(' 0', ' ').replace('/0', '/') + ' - ' + b.strftime(
        "%I:%M%p CST").lstrip('0'))

print '\n'.join(output)
