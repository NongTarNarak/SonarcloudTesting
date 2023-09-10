
DictCharToNum = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 
                 'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12, 
                 'M': 13, 'N': 14, 'O': 15, 'P': 16, 'Q': 17, 'R': 18, 
                 'S': 19, 'T': 20, 'T': 21, 'V': 22, 'W': 23, 'X': 24, 'Y': 25, 'Z': 26, ' ':' '}
listAZ = ['','A', 'B', 'C', 'D', 'E', 'F', 'G', 'H',
              'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
              'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

# Print excel
def DisplayFormula(ArrayTiWantToPrint):
    column100 = len(ArrayTiWantToPrint)
    numAZ = len(ArrayTiWantToPrint[0])

    for i in range(column100+1):
        for j in range(numAZ+1):
            if i == 0 and j==0:
                print('_',end=',')
            elif i == 0 and j!=numAZ:
                print(listAZ[j],end=',')
            elif i == 0 and j==numAZ:
                print(listAZ[j])
                
            if i>0:
                if j == 0 :
                    print(i,end=',')
                elif j>0 and j!=numAZ:
                    print(ArrayTiWantToPrint[i-1][j-1],end=',')
                elif j>0 and j==numAZ:
                    print(ArrayTiWantToPrint[i-1][j-1])
    return ''

# Build None 2D Array
def ArrayBuilder(rowAZ,column):
    numAZ = DictCharToNum[rowAZ]
    NoneArr = []
    for i in range(column):
        x = []
        for i in range(numAZ):
            x.append(' ')
        NoneArr.append(x)
    return NoneArr



def InsertChangeNongTarExcel(arr):
            
    while True:
        pointer = input("** WHICH BOX YOU WANT insert/change Ex. A1 G4 H5 | '0' : STOP INSERT,CHANGE ** ==> ".upper()).upper()
        
        try:
            if(pointer == '0'):
                print()
                return
        
            print()
            print("                      *** INSERT THE VALUE ***")
            print("         ** For Example : Number, '=A8-F4', '=(A2+55)+1', '=5+2' **".upper())
            print("FYI :\n[Cannot use reference as a value. Ex. 'A1', 'G4']\n[Cannot calculate more than two values in parentheses if you use reference Ex. Do : '=(5+3/4)+5-3/4', '=5-4/3*4' | Don't : '5+A2/(4*5-3)']\n[cannot calculate more than 2 values if you use reference].".upper())
            theinputforArr = input(" ==> ").upper()
            if (not checkParenthesis(theinputforArr)):
                continue
            else:
                if  (not theinputforArr.isalpha()):
                    i = int(pointer[1])-1
                    j = DictCharToNum[pointer[0]]-1
                    arr[i][j] = theinputforArr
                    return arr
        
                else:
                    print()
                    print('** Can not input only Character ...Try again **')
                    print()
                    continue
        except:
            print()
            print("** INVALID INDEX ...Try again **")
            print()
            continue
          

def checkParenthesis(myStr):
    stack = []
    for i in myStr:
        if i == '(':
            stack.append(i)
        elif i in ')':
            if ((len(stack) > 0) and ('(' == stack[len(stack)-1])):
                stack.pop()
            else:
                return False
    if len(stack) == 0:
        return True
    else:
        print()
        print("** Error Invalid parentheses ...Try again **".upper())
        print()
        return False


def ForloopError(list):
    for i in range(len(list)):
        for j in range(len(list[0])):
            if('=' in list[i][j] ):
                list[i][j] = '#ERROR'
    return list
# ===================================================Calculation Area=================================================== #

    
def isDigit(x):
    try:
        float(x)
        return True
    except:
        return False
    
from collections import deque
# Expression Tree
class ExprTree:
    def __init__(self, data=None, left=None, right=None):
        self.data = data
        self.left = left
        self.right = right

def calculation(root):
        
        op = ["+", "-", "*", "/"]
        if root is None:
            return
        elif root.data in op:
            if root.left:
                root.left = calculation(root.left)
            if root.right:
                root.right = calculation(root.right)
                
            if root.data == "+":
                return root.left + root.right
            elif root.data == "-":
                return root.left - root.right
            elif root.data == "*":
                return root.left * root.right
            elif root.data == "/":
                return root.left / root.right
        else:

            if (root.left is None and root.right is None):
                return root.data
        return root   

def build_expr_tree(expr_str):
        expr = expr_str.split()
        stack = deque()

        # build expression tree
        for term in expr:
            if term in "+-*/":
                node = ExprTree(term, stack.pop(), stack.pop())
            else:
                node = ExprTree(float(term))

            stack.append(node)

        return stack.pop()
    
def CheckCanCalculateMai(StringX):
    
    countOP = countOperation(StringX)

    if countOP > 1 or countOP <= 0:
        return False
    else:
        if '+' in StringX:
            listOfValues = StringX.split('+')
            listOfValues.insert(0,'+')
        elif '-' in StringX:
            listOfValues = StringX.split('-')
            listOfValues.insert(0,'-')
        elif '*' in StringX:
            listOfValues = StringX.split('*')
            listOfValues.insert(0,'*')
        elif '/' in StringX:
            listOfValues = StringX.split('/')
            listOfValues.insert(0,'/')

        listTiNoChongVang = list(filter(lambda a: a != '', listOfValues))   
                
        if len(listTiNoChongVang) == 3:
            return True
        else:
            return False
    
    
def countOperation(StringX):
    countOP = 0
    for i in StringX:
        if i in ['+','-','*','/']:
            countOP +=1
    return countOP
    
    
# To change string values like '=A9-5' to [5,A9,-] ||| '=Left-right' to [right,left,-] this pattern gonna use to make expr tree
# NO ( )
def ChangeStringValuesToListCaseEasy(stringval):

    if '(' in stringval:
        return str(ChangeStringValuesToKumTobBabNumberCaseAdvance(stringval))
        
    stringval = stringval.replace('=','')
    if '+' in stringval:
        listOfValues = stringval.split('+')
        listOfValues.insert(0,'+')
    elif '-' in stringval:
        listOfValues = stringval.split('-')
        listOfValues.insert(0,'-')
    elif '*' in stringval:
        listOfValues = stringval.split('*')
        listOfValues.insert(0,'*')
    elif '/' in stringval:
        listOfValues = stringval.split('/')
        listOfValues.insert(0,'/')
    # Base case to return to update list
    elif stringval == ' ':
        return '0'
    elif isDigit(stringval):
        return stringval

    listOfValuesTiSortForBuildTree = []   
        
    for i in listOfValues:
        listOfValuesTiSortForBuildTree.insert(0,i)

    for i in range(len(listOfValuesTiSortForBuildTree)):
        if(isDigit(listOfValuesTiSortForBuildTree[i])):
            continue
        else:
            if (listOfValuesTiSortForBuildTree[i] in '+-*/'):
                continue
            else:
                x = firstarr[int(listOfValuesTiSortForBuildTree[i][1:len(listOfValuesTiSortForBuildTree[i])])-1][DictCharToNum[listOfValuesTiSortForBuildTree[i][0]]-1]
                listOfValuesTiSortForBuildTree[i] = ChangeStringValuesToListCaseEasy(x)
          
    return listOfValuesTiSortForBuildTree


# Have ( ) =(4-5)*(5/4)+4*2+8
def ChangeStringValuesToKumTobBabNumberCaseAdvance(stringval):
    stringval = stringval.replace('=','')
    StrX = ''
    for i in stringval:
        if(i != '(' and i !=')'):
            StrX += i
        else:
            StrX += '['
    
    listX = StrX.split('[')
    listTiNoWongLeft = list(filter(lambda a: a != '', listX))

    for i in range(len(listTiNoWongLeft)):
        if listTiNoWongLeft[i] in ['+','-','*','/']:
            continue
        else:
            if(CheckCanCalculateMai(listTiNoWongLeft[i])):  
                listTiNoWongLeft[i] = str(calculation(build_expr_tree(StrBuilderForBuild_expr_tree(listTiNoWongLeft[i])))) #jer simple case talod

    
    strToKumNoanKumTob =''
    for i in listTiNoWongLeft:
        strToKumNoanKumTob += i

    # Now we gonna Make all A1 B1 H3 be the num for the last step to calculate
    strToKumNoanKumTob = strToKumNoanKumTob.replace('.0','')
    
    tempstr01 = strToKumNoanKumTob
    for i in range(len(tempstr01)):
        
        strToDuengInfoFromArr = ''
        
        if(tempstr01[i] in ['+','-','*','/']):
            continue
        elif isDigit(tempstr01[i]):
            continue
        elif tempstr01[i] == '.':
            continue
        else:
            strToDuengInfoFromArr += tempstr01[i]
            for j in range(i+1,len(tempstr01)):
                if isDigit(tempstr01[j]):
                    strToDuengInfoFromArr += tempstr01[j]
                else:
                    break

            # PaiDeuengInfo ma jark list if digit replace to str dai else: kumnoan recursive
            x = firstarr[int(strToDuengInfoFromArr[1:len(strToDuengInfoFromArr)])-1][DictCharToNum[strToDuengInfoFromArr[0]]-1]

            if isDigit(x):
                strToKumNoanKumTob = strToKumNoanKumTob.replace(strToDuengInfoFromArr,x)
            elif (x == ' '):
                strToKumNoanKumTob = strToKumNoanKumTob.replace(strToDuengInfoFromArr,'0')
            elif ('(' in x): 
                PaiKumNoanMaBabAdvanceCase = str(ChangeStringValuesToKumTobBabNumberCaseAdvance(x))
                strToKumNoanKumTob = strToKumNoanKumTob.replace(strToDuengInfoFromArr, PaiKumNoanMaBabAdvanceCase)
            else:   
                PaiKumNoanMaBabEasyCase = str(calculation(build_expr_tree(StrBuilderForBuild_expr_tree(listTiNoWongLeft[i]))))
                strToKumNoanKumTob = strToKumNoanKumTob.replace(strToDuengInfoFromArr,PaiKumNoanMaBabEasyCase)

    return eval(strToKumNoanKumTob) #calculate all str '+-*/'
    
        
# input x is String of values in arr  
def StrBuilderForBuild_expr_tree(stringval):
    # make ['0', ['2', ['5', '2', '*'], '+'], '+'] to '0 2 5 2 * + +'
    def dothing(listComputed):
        strlol = []
        def visit_list_rec(x):
            # base case, we found a non-list element
            if type(x) != list:
                strlol.append(x)
                return

            # recursive step, visit each child
            for element in x:
                visit_list_rec(element)
        # Call method to make traversal all element
        visit_list_rec(listComputed)
        
        donestr = ''
        for i in range(len(strlol)):
            if(i == len(strlol)-1):
                donestr += strlol[i]
            else:
                donestr += strlol[i] + ' '
        return donestr
    
    # ex =A8+5
    listOfValuesTiSortForBuildTree = ChangeStringValuesToListCaseEasy(stringval)
    return dothing(listOfValuesTiSortForBuildTree)

def CalculateAllFirstArr(firstarr,secondarr):
    for i in range(len(firstarr)):
        for j in range(len(firstarr[0])):
            if '=' in firstarr[i][j]:
                try:
                    num = firstarr[i][j].replace('=','')
                    secondarr[i][j] = eval(num)
                except:
                    if '(' in firstarr[i][j]:
                        kumtob = ChangeStringValuesToKumTobBabNumberCaseAdvance(firstarr[i][j])
                    else:
                        kumtob = calculation(build_expr_tree(StrBuilderForBuild_expr_tree(firstarr[i][j])))
                        
                    secondarr[i][j] = kumtob
            elif isDigit(firstarr[i][j]):
                secondarr[i][j] = firstarr[i][j]
                        

# ===================================================Calculation Area=================================================== #

# Main
if __name__ == "__main__":
    # this try is for unexpected exception
    while True:
        try:
            print("** WELCOME TO PAKKAPHAN's EXCEL **")
            print("** FIRST YOU NEED TO CREATE THE EMPTY EXCEL **")
            
            # We use .upper() to debug sensitive case
            while True:
                rowAZ = input("Please insert A-Z : ".upper()).upper()
                column100 = input("Please insert 1-100 : ".upper())
                if(rowAZ in listAZ and isDigit(column100)):
                    # if it all correct break and initiate the array
                    if (int(column100)<=100 and int(column100) >=1):
                        break
                print()
                print('** INVALID INPUT TO INITIATE THE EXCEL **')
                print()
            
            
            # Initiate the array the first one is to save all the first input ex '=A5+5' '4'
            # the second one is for calculate ,,,,, update from the first array
            
            firstarr =  ArrayBuilder(rowAZ,int(column100))
            secondarr = ArrayBuilder(rowAZ,int(column100))
            
            print()
            print(DisplayFormula(firstarr))
            print()
            while True:
                command01 = int(input("( '0' : EXIT | '1' : INSERT/CHANGE EXCEL *Auto display* ) ==> "))
                if(command01 == 0):
                    print()
                    print("** GOOD BYE **")
                    break
                elif(command01 == 1):
                    # For blocking Circular reference
                    try:
                        InsertChangeNongTarExcel(firstarr)
                        print('\n** INSERT COMPLETE **\n')
                        CalculateAllFirstArr(firstarr,secondarr)
                        print('\nFirst array :')
                        DisplayFormula(firstarr)
                        print('\nSecond array :')
                        DisplayFormula(secondarr)
                        print()
                        continue
                    except:
                        ForloopError(firstarr)
                        DisplayFormula(firstarr)
                        print()
                        print("*** The Program is ended due to Circular reference / invalid reference ***".upper())
                        break #stop the program

        except:
            print()
            print("*** There is unexpected error / '0' : Exit | 'type anything' : reboot ***".upper())
            x = input("===> ")
            if(x == '0'):
                break
            else:
                continue
        print()
        break
                
