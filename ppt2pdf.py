# C:\Python27\ArcGIS10.6\python.exe ppt2pdf.py

import sys, os, glob, win32com.client, datetime, time

# start = os.path.dirname(os.path.dirname(__file__))

def convert(files, formatType = 32):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    for filename in files:
        newname = os.path.splitext(filename)[0] + ".pdf"
        deck = powerpoint.Presentations.Open(filename)
        deck.SaveAs(newname, formatType)
        deck.Close()
    powerpoint.Quit()

def main():
    print("\n    Start Time : " + datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))
    input = raw_input('   DRAG the folder conatining your ppts here -->')
    for root, dirs, files in os.walk(input):
        for file in files:
            if file.endswith(".pptx"):
                print(os.path.join("C:\\GitHub\\python\\ppt2pdf",file))
                f = glob.glob(os.path.join(root, file)) # <--- ONLY CHANGE
                convert(f)

    print("\n    End Time : " +  str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " \n\n    COMPLETED!")

if __name__ == '__main__':
    main()
