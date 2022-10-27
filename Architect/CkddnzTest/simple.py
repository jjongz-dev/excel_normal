

languages = ['python', 'perl', 'c', 'java']

for lang in languages:
    if lang in ['python', 'perl']:
        print("%6s need interpreter" % lang)
    elif lang in ['c', 'java']:
        print("%6s need compiler" % lang)
    else:
        print("shouild not reach here")



windows_dict = {}
windows_dict['aaa'] = [1,2,3,4,5]
windows_dict['bbb'] = [6,7,8,9,10]
windows_dict['name'] = 'pey'
windows_dict['3'] = [1,2,3]
del windows_dict['name']

dictt = {"김연아":"피겨스케이팅", "류현진":"야구", "박지성":"축구", "귀도":"파이썬"}


dictt['김연아']




print(dictt['김연아'])
