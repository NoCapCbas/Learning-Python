import pytube
link = 'https://www.youtube.com/watch?v=0sMtoedWaf0'
# poor quality video
# video_download = pytube.YouTube(link)
# video_download.streams.first().download()

# high quality video
video_download = pytube.YouTube(link)
video_download.streams.filter(progressive=True).last().download()
