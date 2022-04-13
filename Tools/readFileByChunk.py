pathToTotal = r"Y:\_PyScripts\Damon\Log\usCensus\totalCENSUS.txt"




def readInChunks(fileObj, chunkSize=2048):
    """
    Lazy function to read a file piece by piece.
    Default chunk size: 2kB.

    """
    while True:
        data = fileObj.read(chunkSize)
        if not data:
            break
        yield data

f = open(pathToTotal)
count = 0
for chunk in readInChunks(f):
    # if count == 0:
    #     print(chunk)
    # count += 1
    print(chunk)
f.close()


readInChunks(pathToTotal)
