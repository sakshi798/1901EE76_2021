import os
os.system("cls")


"""  TASK-1:
     Function to make a file for every individual ROLL NUMBER 
    listed in the "regtable old.csv". All the outputs relating
     to Task 1 will go into "output_individual_roll" folder.
"""

def output_individual_roll():

    """Creating folder output_individual_roll for storing outputs
       for task 1.
    """
    os.mkdir("./output_individual_roll")

    """loop_no2 to check if it's the first iteration in the
        file , we skip the first iteration or first line
        of the file containing variable names/labels
    """

    loop_no2=0

    """reading the file regtable_old.csv"""

    with open('regtable_old.csv', 'r') as fr:
        
        """reading the file line by line"""

        for line in fr:

            """if first iteration then skip it"""

            loop_no2+=1
            if loop_no2==1:
                continue

            """split the line into words separated using ',' separator"""

            words = line.split(',')

            """if the the file does not exists then create file 
               and add headers as the first line
            """

            if os.path.isfile("./output_individual_roll/"+words[0]+".csv")==False:
                with open("./output_individual_roll/"+words[0]+".csv", "w") as fr1:
                    fr1.write("rollno,register_sem,subno,sub_type\n")
                    fr1.close()

            """adding the data i.e, the details related to the roll number"""

            fr2 = open("./output_individual_roll/"+words[0]+".csv", "a")
            fr2.write(words[0] + "," + words[1] + "," + words[3] + "," +words[8])
            fr2.close()
        fr.close()
    return



"""  TASK-2:
     Function to make a file for every individual subject 
    listed in the "regtable old.csv". All the outputs relating
     to Task 2 will go into "output_by_subject" folder.
"""

def output_by_subject():

    """Creating folder output_by_subject for storing outputs
       for task 2.
    """
    os.mkdir("./output_by_subject")

    """loop_no to check if it's the first iteration in the
        file , we skip the first iteration or first line
        of the file containing variable names/labels
    """
    loop_no=0

    """reading the file regtable_old.csv"""

    with open('regtable_old.csv', 'r') as f:

        """reading the file line by line"""

        for line in f:

            """if first iteration then skip it"""

            loop_no+=1
            if loop_no==1:
                continue

            """split the line into words separated using ',' separator"""

            words = line.split(',')

            """if the the file does not exists then create file 
               and add headers as the first line
            """

            if os.path.isfile("./output_by_subject/"+words[3]+".csv")==False:
                with open("./output_by_subject/"+words[3]+".csv", "w") as f1:
                    f1.write("rollno,register_sem,subno,sub_type\n")
                    f1.close()

            """adding the data i.e, the details related to the subject"""

            f2 = open("./output_by_subject/"+words[3]+".csv", "a")
            f2.write(words[0] + "," + words[1] + "," + words[3] + "," +words[8])
            f2.close()
        f.close()
    return
   
# TASK : 1
output_individual_roll()
# TASK : 2
output_by_subject()


