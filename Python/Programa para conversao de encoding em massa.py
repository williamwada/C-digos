def correctSubtitleEncoding(filename, newFilename, encoding_from, encoding_to='UTF-8'):
    with open(filename, 'r', encoding=encoding_from) as fr:
        with open(newFilename, 'w', encoding=encoding_to) as fw:
            for line in fr:
                fw.write(line[:-1] + '\r\n')


for year in range(1945,2019):
    print("ano: "+str(year))
    correctSubtitleEncoding("proposicoes-"+str(year)+".json", "converted\proposicoes-"+str(year)+"-unicode.json", 'UTF-8', encoding_to='utf-16')