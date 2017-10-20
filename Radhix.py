import speech_recognition as sr
import win32com.client as wincl
import webbrowser
import time
import wolframalpha
import wikipedia
import os
import pyperclip
import win32com.client

banner = """

██████╗  █████╗ ██████╗ ██╗  ██╗██╗██╗  ██╗     █████╗ ██╗
██╔══██╗██╔══██╗██╔══██╗██║  ██║██║╚██╗██╔╝    ██╔══██╗██║
██████╔╝███████║██║  ██║███████║██║ ╚███╔╝     ███████║██║
██╔══██╗██╔══██║██║  ██║██╔══██║██║ ██╔██╗     ██╔══██║██║
██║  ██║██║  ██║██████╔╝██║  ██║██║██╔╝ ██╗    ██║  ██║██║
╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝ ╚═╝  ╚═╝╚═╝╚═╝  ╚═╝    ╚═╝  ╚═╝╚═╝ Ver.1
                                                          """
print (banner)

cl = wolframalpha.Client('7PHER2-83UYUKUL52')
shell = wincl.Dispatch("WScript.Shell")
speak = wincl.Dispatch("SAPI.SpVoice")
r = sr.Recognizer()
r.pause_threshold = 0.7                                                                     #it works with 1.2 as well
r.energy_threshold = 400

print ("Hello! Welcome To Radhix_AI V.1 \nFor a list of commands, plese say 'keyword list'...'")
speak.Speak('Hello! Welcome To Radhix AI Verison One, For a list of commands, plese say "keyword list"...')


#List of Available Commands

keywd = 'keyword list'
google = 'search for'
acad = 'academic search'
sc = 'deep search'
wkp = 'wiki page for'
rdds = 'read this text'
sav = 'save this text'
bkmk = 'bookmark this page'
vid = 'video for'
wtis = 'what is'
wtar = 'what are'
whis = 'who is'
whws = 'who was'
when = 'when'
where = 'where'
how = 'how'
paint = 'open paint'
lsp = 'silence please'
lsc = 'resume listening'
stoplst = 'stop listening'


# main loop
while True:

    with sr.Microphone() as source:  #instantiating the Microphone
        
     
        try:
            print ('listening....!')
            audio = r.listen(source)
            message = str(r.recognize_google(audio))
            print ('Done!')
            print('Radhix_AI.V1 thinks you said:\n' + message)
            speak.Speak(r.recognize_google(audio))
     
    
            if google in message:                                                           #what happens when google keyword is recognized

                words = message.split()
                del words[0:2]
                st = ' '.join(words)
                print('Google Results for: '+str(st))
                url='http://google.com/search?q='+st
                webbrowser.open(url)
                speak.Speak('Google Results for: '+str(st))

            elif acad in message:                                                           #what happens when acad keyword is recognized

                words = message.split()
                del words[0:2]
                st = ' '.join(words)
                print('Academic Results for: '+str(st))
                url='https://scholar.google.ro/scholar?q='+st
                webbrowser.open(url)
                speak.Speak('Academic Results for: '+str(st))

            elif wkp in message:                                                            #what happens when wkp keyword is recognized

                try:

                    words = message.split()
                    del words[0:3]
                    st = ' '.join(words)
                    wkpres = wikipedia.summary(st, sentences=2)

                    try:

                        print('\n' + str(wkpres) + '\n')
                        speak.Speak(wkpres)

                    except UnicodeEncodeError:
                        speak.Speak(wkpres)

                except wikipedia.exceptions.DisambiguationError as e:
                    print (e.options)
                    speak.Speak("Too many results for this keyword. Please be more specific and try again")
                    continue

                except wikipedia.exceptions.PageError as e:
                    print('The page does not exist')
                    speak.Speak('The page does not exist')
                    continue

            elif sc in message:                                                             #what happens when sc keyword is recognized

                try:
                    words = message.split()
                    del words[0:1]
                    st = ' '.join(words)
                    scq = cl.query(st)
                    sca = next(scq.results).text
                    print('The answer is: '+str(sca))
                    #url='http://www.wolframalpha.com/input/?i='+st
                    #webbrowser.open(url)
                    speak.Speak('The answer is: '+str(sca))

                except StopIteration:
                    print('Your question is ambiguous. Please try again!')
                    speak.Speak('Your question is ambiguous. Please try again!')

                else:
                    print('No query provided')

            elif paint in message:                                                          #what happens when paint keyword is recognized
                os.system('mspaint')

            elif rdds in message:                                                           #what happens when rdds keyword is recognized
                print("Reading your text")
                speak.Speak(pyperclip.paste())

            elif sav in message:                                                            #what happens when sav keyword is recognized
                with open('path to your text file', 'a') as f:
                    f.write(pyperclip.paste())
                print("Saving your text to file")
                speak.Speak("Saving your text to file")

            elif bkmk in message:                                                           #what happens when bkmk keyword is recognized
                shell.SendKeys("^d")
                speak.Speak("Page bookmarked")

            elif keywd in message:                                                          #what happens when keywd keyword is recognized

                print ('')
                print ("Say 'search for' to return a Google search")
                print ("Say 'academic search' to return a Google Scholar search")
                print ("Say 'deep search' to return a Wolfram Alpha query")
                print ("Say 'wiki page for' to return a Wikipedia page")
                print ("Say 'read this text' to read the text you have highlighted and Ctrl+C (copied to clipboard)")
                print ("Say 'save this text' to save the text you have highlighted and Ctrl+C-ed (copied to clipboard) to a file")
                print ("Say 'bookmark this page' to bookmark the page your are currently reading in your browser")
                print ("Say 'video for' to return video results for your query")
                print ("Say ' 'stop listening' ' to shut down")
                print ("Say 'open paint' to open paint in windows")
                print ("Say 'silence please' to pause listening")
                print ("Say 'resume listening' to resume listening")
                print ("For more general questions, ask them naturally and I will do my best to find a good answer")

            elif vid in message:                                                            #what happens when vid keyword is recognized

                words = message.split()
                del words[0:2]
                st = ' '.join(words)
                print('Video Results for: '+str(st))
                url='https://www.youtube.com/results?search_query='+st
                webbrowser.open(url)
                speak.Speak('Video Results for: '+str(st))

            elif wtis in message:                                                           #what happens when wtis keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('The answer is: '+str(sca))
                    #url='http://www.wolframalpha.com/input/?i='+st
                    #webbrowser.open(url)
                    speak.Speak('The answer is: '+str(sca))

                except UnicodeEncodeError:

                    speak.Speak('The answer is: '+str(sca))

                except StopIteration:

                    words = message.split()
                    del words[0:2]
                    st = ' '.join(words)
                    print('Google Results for: '+str(st))
                    url='http://google.com/search?q='+st
                    webbrowser.open(url)
                    speak.Speak('Google Results for: '+str(st))

            elif wtar in message:                                                           #what happens when wtar keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('The answer is: '+str(sca))
                    #url='http://www.wolframalpha.com/input/?i='+st
                    #webbrowser.open(url)
                    speak.Speak('The answer is: '+str(sca))

                except UnicodeEncodeError:

                    speak.Speak('The answer is: '+str(sca))

                except StopIteration:

                    words = message.split()
                    del words[0:2]
                    st = ' '.join(words)
                    print('Google Results for: '+str(st))
                    url='http://google.com/search?q='+st
                    webbrowser.open(url)
                    speak.Speak('Google Results for: '+str(st))

            elif whis in message:                                                           #what happens when whis keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    speak.Speak('The answer is: '+str(sca))

                except StopIteration:

                    try:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        wkpres = wikipedia.summary(st, sentences=2)
                        print('\n' + str(wkpres) + '\n')
                        speak.Speak(wkpres)

                    except UnicodeEncodeError:

                        speak.Speak(wkpres)

                    except:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        print('Google Results (last exception) for: '+str(st))
                        url='http://google.com/search?q='+st
                        webbrowser.open(url)
                        speak.Speak('Google Results for: '+str(st))

            elif whws in message:                                                           #what happens when whws keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    speak.Speak('The answer is: '+str(sca))

                except StopIteration:

                    try:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        wkpres = wikipedia.summary(st, sentences=2)
                        print('\n' + str(wkpres) + '\n')
                        speak.Speak(wkpres)

                    except UnicodeEncodeError:

                        speak.Speak(wkpres)

                    except:

                        words = message.split()
                        del words[0:2]
                        st = ' '.join(words)
                        print('Google Results for: '+str(st))
                        url='http://google.com/search?q='+st
                        webbrowser.open(url)
                        speak.Speak('Google Results for: '+str(st))

            elif when in message:                                                         #what happens when 'when' keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    speak.Speak('The answer is: '+str(sca))

                except UnicodeEncodeError:

                    speak.Speak('The answer is: '+str(sca))

                except:

                    print('Google Results for: '+str(message))
                    url='http://google.com/search?q='+str(message)
                    webbrowser.open(url)
                    speak.Speak('Google Results for: '+str(message))

            elif where in message:                                                        #what happens when 'where' keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    speak.Speak('The answer is: '+str(sca))

                except UnicodeEncodeError:

                    speak.Speak('The answer is: '+str(sca))

                except:

                    print('Google Results for: '+str(message))
                    url='http://google.com/search?q='+str(message)
                    webbrowser.open(url)
                    speak.Speak('Google Results for: '+str(message))

            elif how in message:                                                          #what happens when 'how' keyword is recognized

                try:

                    scq = cl.query(message)
                    sca = next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    speak.Speak('The answer is: '+str(sca))

                except UnicodeEncodeError:

                    speak.Speak('The answer is: '+str(sca))

                except:

                    print('Google Results for: '+str(message))
                    url='http://google.com/search?q='+str(message)
                    webbrowser.open(url)
                    speak.Speak('Google Results for: '+str(message))

            elif stoplst in message:                                                        #what happens when stoplst keyword is recognized
                speak.Speak("I am shutting down")
                print("Shutting down...")
                break

            elif lsp in message:

                speak.Speak('Listening is paused')
                print('Listening is paused')
                r2 = sr.Recognizer()
                r2.pause_threshold = 0.7
                r2.energy_threshold = 400

                while True:

                    with sr.Microphone() as source2:

                        try:

                            audio2 = r2.listen(source2, timeout = None)
                            message2 = str(r.recognize_google(audio2))

                            if lsc in message2:
                                speak.Speak('I am listening')
                                break

                            else:
                                continue

                        except sr.UnknownValueError:
                            print("Listening is paused. Say resume listening when you're ready...")

                        except sr.RequestError:
                            speak.Speak("I'm sorry, I couldn't reach google")
                            print("I'm sorry, I couldn't reach google")


            else:
                pass

        except sr.UnknownValueError:
            print("For a list of commands, say: 'keyword list'...")

        except sr.RequestError:
            speak.Speak("I'm sorry, I couldn't reach google")
            print("I'm sorry, I couldn't reach google")

    time.sleep(0.3)
