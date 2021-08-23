"""Function to check if the given input list is valid"""
def is_Valid(input_nums):
    """false value of flag indicates invalid input list"""
    flag=False
    """List to store invalid inputs"""
    invalid_inputs=[]

    for no in input_nums:
        x=str(no)
        if(x.isdigit()):
            continue
        else:
            """If the no is not a digit (i.e, integer) , then change 
            the value of flag to true indicating invalid input"""
            flag=True
            """adding the invalid input to the list invalid_inputs"""
            invalid_inputs.append(no)

    """If flag is true means that there exists atleast one invalid input"""
    if(flag):
        print("Please enter a valid input list. Invalid inputs detected:")
        print(invalid_inputs)
    return not(flag)

"""Function to calculate the score in the memory game"""
def get_memory_score(input_nums):
    """initialising score with zero"""
    score=0
    """list to store numbers present in players memory"""
    player_memory=[]

    """traversing through each number in list"""
    for i in range(len(input_nums)):
        
        while(len(player_memory)>5):
            player_memory.pop(0)
        """if there are more tha 5 numbers in the list
        remove the first number in the list"""
        
        flag=True

        """ If the called number is already in the player’s memory, 
        a point is added to the player’s score"""
        for x in range(len(player_memory)):
            if(player_memory[x]==input_nums[i]):
                score+=1
                flag=False
                break
            x+=1

        """ If the called number is NOT present in the player’s memory, 
            add it to the player’s memory"""
        if flag:
            player_memory.append(input_nums[i])

    return score    


input_nums =  [3, 4, 3, 0, 7, 4, 5, 2, 1, 3]

"""if the given input list is valid calculate the score
and print the score as the output"""
if (is_Valid(input_nums)):
    print("Score : {score}".format(score=get_memory_score(input_nums)))
