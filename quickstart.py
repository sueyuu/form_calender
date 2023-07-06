#!/usr/bin/env python
# coding: utf-8

# In[6]:


import calendar
from datetime import datetime, date, time, timedelta

from dateutil.parser import parse, parserinfo

import re
from random import randrange
import pandas as pd

import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request


# In[7]:


def get_eve():
  eve_list = []

  now = datetime.now()
  yr = now.year
  mn = now.month
  s = 1
  _, e = calendar.monthrange(yr, mn)
# If modifying these scopes, delete the file token.json.
  SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

  creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
  if os.path.exists('token.json'):
      creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
      creds = flow.run_local_server(port=0)
# Save the credentials for the next run
      with open('token.json', 'w') as token:
        token.write(creds.to_json())
  service = build('calendar', 'v3', credentials=creds)


  page_token = None

  eventlist = service.events().list(calendarId= 'dmclinic.yu@gmail.com', 
                                  timeMin = datetime(yr,mn,s).isoformat()+ 'Z',  
                                  timeMax = datetime(yr,mn,e).isoformat()+ 'Z', 
                                  singleEvents=True).execute()

  for event in eventlist['items']:
    try:
      eve_list.append((event['summary'], event['start']['date'], event['end']['date']))
    except:

      try:
        eve_list.append((event['summary'], event['start']['dateTime'], event['end']['dateTime']))
      except:
        pass
  return eve_list


# In[18]:


class CustomParserInfo(parserinfo):
  JUMP= [' ', '.', ',', ';', '-', '/', "'", 'at', 'on', 'and', 'ad', 'm', 't', 'of', 'st', 'nd', 'rd', 'th']+['點']
  @staticmethod
  def get_jump():
    return CustomParserInfo.JUMP


def managed_evt(eve_list, sentence_sq):
    
  for eve, s, e in eve_list:
#按序search
    for sq in sentence_sq:
      copy_eve = eve
      for seg in sq:

        for wd in seg:
#找到後應該下一次要從斷點開始找
 
          if re.search(wd, copy_eve):

            init_point, new_start_point = re.search(wd, copy_eve).span()
#modified_eve從始錨點到終錨點結束
            if seg == sq[0]:
              modified_eve = copy_eve[init_point:new_start_point]

            else:              
              modified_eve = modified_eve+copy_eve[:new_start_point+1]              

            copy_eve = copy_eve[new_start_point:]
#keep on 'seg in sq' loop with new eve if word found
            break
            
        else:
#if word not found, break this sq, move to next sq
#modified_eve is None in this sq
          modified_eve = None
          break
        
#if nickname and off, try find timeperiod or dateparser
        if seg == sq[-1] and sq == [nickname, off]:
#先找時段，如找到不用parse，change to key

          for period, key_word in time_period.items():
        
            modified_eve = re.sub(key_word, period, modified_eve)            
              
#if no timeperiod, try parse modified_eve
          if not re.search(r'morning|afternoon', modified_eve):

            try:
##if parse modified_eve success
              if parse(modified_eve, parserinfo=CustomParserInfo(), fuzzy_with_tokens = True):
                                   
#if parse modified_eve第一次成功, 先找ampm
                if re.search(r'am|pm', modified_eve.lower()):
                  
                  timing_list = ampm_seperator(re.split(r'(am|pm)', modified_eve.lower()))
#回傳的是以ampm當最後一位的sentence list，待會要再做一次parse
                  timing_list = [part for part in timing_list]
#if no ampm, split sentence with JUMP, save to timing_list(sentence list), 再parse一次
                else:

                  jump = CustomParserInfo.get_jump()
                  jump.remove('點')
                  jump.remove(' ')
                  jump.remove('.')
                  jump.append('~')
                  for seperator in jump:
                    if len(modified_eve.split(seperator))>1:
                      timing_list = modified_eve.split(seperator)
                      break
#parse each sentence in timing_list, save to parsed_tuple
                if timing_list:

                  parsed_tuple = []
                  for each_s in timing_list:
                    
                    try:
                      parsed_tuple.append(parse(each_s, parserinfo=CustomParserInfo(), fuzzy_with_tokens = True))                    
                    except:
                      pass
                    finally:
                      continue
#generated parsed_tuple from modified_eve and time from s and e if s and e has time information
#only generate parsed_tuple if time info in s and e(先預設s and e在同天)
              else:
        
                if datetime.fromisoformat(s).time()>time(0,0,0):
                  parsed_tuple = [(datetime.fromisoformat(s), (modified_eve,)), 
                                 (datetime.fromisoformat(e), (modified_eve,))]

            except:
              pass              
#try->else: 預設parsed_tuple not empty(means either date parser ok or from s,e)
            else:
              modified_eve = ''
#get datetime element in parsed_tuple and sort, if > 12:00pm, 下午
              for dt, other_word in parsed_tuple:
                if dt.time() <= time(12,0,0):
                  modified_eve += "".join(other_word[:-1])
                  modified_eve += "".join((other_word[-1], 'morning'))
                else:
                  modified_eve += "".join(other_word[:-1])
                  modified_eve += "".join((other_word[-1], 'afternoon'))

#if modified_eve in this sq, yield and break sq in sentence_sq, move to next eve
        
      if modified_eve:

        yield modified_eve, s, e
        break
        
        
#seperate with keyword ampm

  def ampm_seperator(word_list):
    final_str = ''
    while True:
      new_wd = word_list.pop(0)
      final_str += new_wd
      if new_wd =='am' or new_wd =='pm':
        yield final_str
        final_str =''
      if 'pm' not in word_list and 'am' not in word_list:
        return


# In[9]:


def form_calendar(emp, dep, shift, occations):

  shift_time_list=[(time(7,30,0),time(12,0,0)),
                  (time(14,30,0),time(17,0,0))]

  now = datetime.now()
  yr = now.year
  mn = now.month

  cl_head = ['姓名','部門','刷卡日期','刷卡時間']    
  pd.DataFrame(columns = cl_head).to_csv(f'{yr, mn, emp}.csv', index=False, encoding = 'utf-8-sig')
    
  s = 1
  _, e = calendar.monthrange(yr, mn)
  mn_std = datetime(yr, mn, s)
  mn_end = datetime(yr, mn, e)

  cl = calendar.monthcalendar(yr, mn)
#先用shift 和 cl建立新的calendar, 日期變成[date, shift(on or off), shift(on or off)]
  
  cl = [(d, i) for wk in cl for d, s in zip(wk, shift) for i in range(s) if d and s]
  
  for eve, std, end in occations:
    std=parse(std).replace(tzinfo=None)      
    end=parse(end).replace(tzinfo=None)

    if std<mn_std:
      std = mn_std.day
    else:
      std = std.day
    if end>mn_end:
      end = mn_end.day+1
    else:
      end = end.day
    
    if std == end:
      end = end+1
    
    if re.search('上班日', eve):
      cl.append((std, 0))
    else:
      
      for d in range(std, end):
        if not re.search('morning|afternoon', eve):
          try:
            cl.remove((d, 0))
            cl.remove((d, 1))
          except:
            pass
          continue
        if re.search('morning', eve):
          try:
            cl.remove((d, 0))
          except:
            pass
        if re.search('afternoon', eve):
          try:
            cl.remove((d, 1))
          except:
            pass
  cl = sorted(cl, key = lambda ds_pair:ds_pair[0])  
  for d, s in cl:

    dt_on=datetime.combine(datetime(yr,mn,d),shift_time_list[s][0])-timedelta(minutes=randrange(10), seconds=randrange(59))
    dt_off=datetime.combine(datetime(yr,mn,d),shift_time_list[s][1])+timedelta(minutes=randrange(10),seconds=randrange(59))
    pd.DataFrame([[emp, dep, dt_on.date(), dt_on.time()], [emp, dep, dt_off.date(), dt_off.time()]],
                 columns = cl_head).to_csv(f'{yr, mn, emp}.csv', index=False, header = False, mode = 'a', 
                                           encoding = 'utf-8-sig')


# In[19]:


if __name__ == '__main__':
  nickname = ['駱', '書羽']
  clinic = ['診所']
#目前無夜診，先不處理夜診
  time_period = {'morning':'早|上午', 'afternoon': '^午|[^上]午'}
  festival = ['中秋', '春節', '端午', '清明']
  off = ['休', '假']
  on = ['上班日']
  big_thing = ['大休診']

#排序(目前不包含時間在最前面或時間在最後面的sentence)
  general_sentence_sq = [[festival,off],
  [clinic,off],
  [big_thing]]
    
  my_sentence_sq = general_sentence_sq + [[nickname, off], [nickname, on]]
  eve_list = get_eve()

  my_occations = [ele for ele in managed_evt(eve_list, my_sentence_sq)]

  form_calendar('駱書羽', '游能俊診所', [2,2,2,2,1,0,0], my_occations)


# In[ ]:




