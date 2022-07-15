import os, shutil, zipfile, numpy as np
from PIL import Image, ImageChops

# ----- finding the only Word file in folder
for file in os.listdir():
    if file.endswith('docx'):
        wordfile = file

# ----- extracting it as archive to temporary folder
with zipfile.ZipFile(wordfile, 'r') as zip_ref:
    zip_ref.extractall("targetdir")

# ----- copying all images to folder with watermarks
os.chdir(r'.\targetdir\word\media')
for file in os.listdir():
    shutil.copyfile(file, '..\..\..\html\\' + file)

# ----- changing folder to where watermarks are
os.chdir('..\..\..\html')

# ----- removing links to Dr. Explain
pages = [i for i in os.listdir() if i.endswith('htm')]
for page in pages:
    inp = open(page, 'r', encoding='utf-8')
    lines = inp.readlines()
    output = ''
    flag = True
    for i in lines:
        if 'h6' not in i and 'Unregistered version' not in i and flag:
            output += i
        elif 'h6' in i:
            flag = not flag
            continue
        else:
            output += '</div>'
    inp.close()
    out = open(page, 'w', encoding='utf-8')
    print(output, file=out)
    out.close()
    
# ----- removing watermarks
water = [i for i in os.listdir() if i.startswith('drex') and i.endswith('png') and 'header' not in i and 'index' not in i]
nowater = [i for i in os.listdir() if i.startswith('image') and i.endswith('png')]

cnt = 0

for watered_img in water:
    cnt += 1
    img1 = Image.open(watered_img)
    width, height = img1.width, img1.height
    if width < 120 and height < 120:
        shutil.copyfile(watered_img, watered_img[:-4] + '_0.png')
        continue
    maxdif = 10 ** 9
    for i in nowater:
        img = Image.open(i)
        if img.width < 120 and img.height < 120:
            continue
        img = img.resize((width, height))
        res = ImageChops.difference(img1.convert('RGB'), img.convert('RGB'))
        mean = np.mean(np.array(res))
        if mean < maxdif:
            maxdif = mean
            ans = i
    shutil.copyfile(ans, watered_img[:-4] + '__0.png')
    print('%.2f' % (cnt / len(water) * 100), '% done...', sep='')

# ----- removing watermarks files
for i in os.listdir():
    if i.startswith('drex') and i.endswith('png') and '__0' not in i and 'header' not in i and 'index' not in i:
        os.remove(i)

# ----- restoring images from what's found
for i in os.listdir():
    if i.startswith('drex') and i.endswith('png'):
        os.rename(i, i.replace('__0', ''))

# ----- deleting images from Word
# ----- (some are not removable for some reason)
for i in os.listdir():
    if i.startswith('image') and i.endswith('png'):
        try:
            os.remove(i)
        except:
            print('Could not delete ' + i)

# ----- deleting temporary folder with Word file content
os.chdir(r'..')
shutil.rmtree('targetdir')

a = input('Press Enter to finish...')