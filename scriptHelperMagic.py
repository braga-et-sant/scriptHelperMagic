import os
from shutil import copyfileobj

from docx import Document
from requests import get
from os.path import exists



def docScour(filename, color):
    color = color.replace("(", "").replace(")", "")
    colorI = tuple(map(int, color.split(', ')))
    doc = Document(filename)
    fullText = []
    for para in doc.paragraphs:
        for count, run in enumerate(para.runs):
            currentRun = run
            text2add = ""
            skipped = 1
            while (currentRun.font.color.rgb == colorI):
                text2add += currentRun.text
                currentRun = para.runs[count+1]
                count+=1
            if(len(text2add) > 0):
                fullText.append(text2add)

    #print(fullText)
    return fullText


def card2numb(img):
    baseurl = "https://api.scryfall.com/cards/named?fuzzy="
    first_response = get(baseurl + img, timeout=5)
    response_list = first_response.json()
    #print(response_list)
    #print(response_list['image_uris']['png'])
    if(response_list['object'] == "error"):
        print("Error: Bad Token: " + img)
        return "error"
    else:
        return (response_list['image_uris']['png'], response_list['id'])


if __name__ == '__main__':
    scriptname = input("Script name:")
    scriptwithExt = scriptname + ".docx"
    fnameImg = "img/"
    #fscript = scriptname + "/"
    color = "(255, 0, 0)"
    data = ""
    #imglist = docScour(scriptwithExt, color)
    imglist = docScour(scriptwithExt, color)
    #print(imglist)
    idcheck = []
    if not (os.path.exists(fnameImg)):
        os.mkdir(fnameImg)
    if not (os.path.exists(fnameImg + scriptname + "/")):
        os.mkdir(fnameImg + scriptname + "/")

    for i in imglist:
        try:
            #print(fnameImg + scriptname + "/" + i.lstrip().rstrip() + ".png")
            #print(not (os.path.exists(fnameImg + scriptname + "/" + i.lstrip().rstrip() + ".png")))
            if not (os.path.exists(fnameImg + scriptname + "/" + i.lstrip().rstrip() + ".png")):
                if card2numb(i)[1] in idcheck:
                    pass
                else:
                    print("Downloading " + i.lstrip().rstrip() + "...")
                    url = card2numb(i)[0]
                    response = get(url, stream=True)
                    response.raw.decode_content = True

                    with open(fnameImg + scriptname + "/" + i + ".png", 'wb') as f:
                        copyfileobj(response.raw, f)

                    idcheck.append(card2numb(i)[1])
        except:
            pass



    print("All Done! Feel free to close this window")
    input("")
