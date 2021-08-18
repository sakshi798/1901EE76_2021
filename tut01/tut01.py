# NAME : SAKSHI SINGH
# ROLL NO: 1901EE76
# TUT 01 CODE
"""This function will detect if a number is meraki number"""

def meraki_helper(n):
    x=n

    """ create two variables last and second_last to store two consecutive digits
     we start considering digits from last , and then divide number by 10"""

    last=x%10                   
    x=int(x/10)

    """ If the given number is a single digit number then it is a Meraki number
     variable ans of type bool stores true if given number is a meraki number"""

    if x==0:
        ans=True
    else:
        ans=True

        """ running a while loop to consider all the adjacent pair of digits in the number
             if the difference is not one then number is not a meraki number"""

        while x>0:
            #initialise second_last digit
            second_last=x%10 
            """absolute difference between two consecutive digits should be 1"""
            if (abs(second_last-last)!=1):
                ans=False
                break   
            #update last digit and x
            last=second_last 
            x=int(x/10)

    if ans:
        print("Yes - {} is a Meraki number".format(n))
    else :
        print("No - {} is not a Meraki number".format(n))

    return ans    


"""count_of_Meraki stores count of numbers in the list that are meraki
    and non_Meraki stores count of numbers in the list that are not meraki"""   

count_of_Meraki=0
non_Meraki=0
input = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]

"""For every number in the list find whether it is a Meraki number or not
    if the helper function returns true increment count_of_Meraki otherwise 
    increment non_Meraki"""
for i in input:
    if(meraki_helper(i)):
        count_of_Meraki=count_of_Meraki+1
    else:
        non_Meraki=non_Meraki+1

print("\nThe given input list contains {} meraki and {} non meraki numbers.".format(count_of_Meraki,non_Meraki))
    