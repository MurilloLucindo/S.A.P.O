import base64


def pic2str(file, functionName):
    pic = open(file, 'rb')
    content = '{} = {}\n'.format(functionName, base64.b64encode(pic.read()))
    pic.close()

    with open('pic2str.py', 'a') as f:
        f.write(content)


if __name__ == '__main__':
    pic2str('forg.png', 'forg')