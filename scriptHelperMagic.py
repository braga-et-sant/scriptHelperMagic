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
                try:
                    currentRun = para.runs[count+1]
                except:
                    break
                count+=1
            if(len(text2add) > 0):
                fullText.append(text2add)

    #print(fullText)
    return fullText


def card2numb(img):
    baseurl = "https://api.scryfall.com/cards/named?exact="
    first_response = get(baseurl + img, timeout=5)
    response_list = first_response.json()
    #print(response_list)
    if (response_list['object'] == "error"):
        print("Error: Bad Token: " + img)
        return "error"
    elif 'card_faces' in response_list.keys():
        #print(response_list['card_faces'][0])
        return (response_list['card_faces'][0]['image_uris']['png'] + '%' + response_list['card_faces'][1]['image_uris']['png']), (response_list['id'])
    else:
        return (response_list['image_uris']['png'], response_list['id'])


if __name__ == '__main__':

    scriptname = input("Script name:")
    scriptwithExt = "script/" + scriptname + ".docx"
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
                cardurl, cardid = card2numb(i)
                if cardid in idcheck:
                    pass
                elif '%' in cardurl:
                    cardurl1, cardurl2 = cardurl.split('%')
                    print("Downloading double faced card " + i.lstrip().rstrip() + "...")

                    response = get(cardurl1, stream=True)
                    response.raw.decode_content = True

                    with open(fnameImg + scriptname + "/" + i + " Front.png", 'wb') as f:
                        copyfileobj(response.raw, f)

                    response2 = get(cardurl2, stream=True)
                    response2.raw.decode_content = True

                    with open(fnameImg + scriptname + "/" + i + " Back.png", 'wb') as f:
                        copyfileobj(response2.raw, f)

                    idcheck.append(cardid)

                else:
                    print("Downloading " + i.lstrip().rstrip() + "...")
                    response = get(cardurl, stream=True)
                    response.raw.decode_content = True

                    with open(fnameImg + scriptname + "/" + i + ".png", 'wb') as f:
                        copyfileobj(response.raw, f)

                    idcheck.append(cardid)
        except:
            pass



    print("All Done! Feel free to close this window")
    input("")
