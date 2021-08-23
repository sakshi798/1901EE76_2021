def is_Valid(input_nums):
    flag=False
    invalid_inputs=[]

    for no in input_nums:
        x=str(no)
        if(x.isdigit()):
            continue
        else:
            flag=True
            invalid_inputs.append(no)

    if(flag):
        print("Please enter a valid input list. Invalid inputs detected:")
        print(invalid_inputs)
    return not(flag)


def get_memory_score(input_nums):
    score=0
    j=0
    list_new=[]

    for i in range(len(input_nums)):
        while(len(list_new)>5):
            list_new.pop(0)

        x=0
        flag=True

        while(x<len(list_new)):
            if(list_new[x]==input_nums[i]):
                score+=1
                flag=False
                break
            x+=1

        if flag:
            list_new.append(input_nums[i])

    return score    


input_nums =  [3, 4, 3, 0, 7, 4, 5, 2, 1, 3]

if (is_Valid(input_nums)):
    print("Score : {score}".format(score=get_memory_score(input_nums)))
