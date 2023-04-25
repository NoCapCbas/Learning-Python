import pytube
link = 'https://www.youtube.com/watch?v=0sMtoedWaf0'
video_download = pytube.Youtube(link)
video_download.streams.first().download()
video_download = pytube.Youtube(link)
video_download.streams.filter(progressive=True).last().download()
