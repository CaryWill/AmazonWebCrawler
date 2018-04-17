pagenumber = 0
def exceptionTest(pagenumber):
    try:
        
    except:
        print(pagenumber)
        exceptionTest(pagenumber)
def main():
    for i in range(1,10):
        exceptionTest(i)