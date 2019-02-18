import datetime
import json
import sys
import xlsxwriter
import zipfile

def parse_users(users_file):
    users = {}
    for user in json.loads(users_file.read()):
        display_name = user.get('real_name', user['name'])
        users[user['id']] = {
            'display_name': display_name,
            'replace_args': ('<@%s>' % user['id'], '@%s' % display_name),
        }
    return users
    
def parse_json(json_file, data, users):
    json_contents = json.loads(json_file.read())
    for message in json_contents:
        if message.get('subtype') == 'channel_join': continue
        if 'attachments' in message: continue # TODO(ronshapiro)

        thread_key = message.get('thread_ts', message['ts'])
        thread_data = data.setdefault(thread_key, [])
        
        user_id = message['user']
        if user_id in users:
            display_name = users[user_id]['display_name']
        else:
            display_name = user_id

        text = message['text']
        for user in users.itervalues():
            text = text.replace(*user['replace_args'])

        thread_data.append({'user': display_name,
                            'message': text,
                            'timestamp': datetime.datetime.fromtimestamp(float(message['ts'])),
        })


with zipfile.ZipFile(sys.argv[1], 'r') as slack_archive:
    all_data = {}
    users = parse_users(slack_archive.open('users.json', 'r'))
    
    for f in slack_archive.namelist():
        slash_index = f.find('/')
        if slash_index != -1 and f.endswith('.json'):
            sheet_name = f[0:slash_index]
            sheet_data = all_data.setdefault(sheet_name, {})
            parse_json(slack_archive.open(f, 'r'), sheet_data, users)

    workbook = xlsxwriter.Workbook('slack_archive.xlsx',
                                   {'default_date_format': 'mmm d yyyy H:mm AM/PM UTC'})
    bold = workbook.add_format({'bold': True})
    justify_top = workbook.add_format({'valign': 'top'})
   
    channels = list(all_data.keys())
    channels.sort()
    for channel in channels:
        worksheet = workbook.add_worksheet('#' + channel)
        worksheet.write_row(0, 0, ['Date', 'Name', 'Text', 'Reply Text'], bold)
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 1, 18)
        
        next_row = 1
        keys = list(all_data[channel].keys())
        keys.sort()
        for key in keys:
            messages = all_data[channel][key]
            worksheet.write(next_row, 0, messages[0]['timestamp'])
            worksheet.write(next_row, 1, messages[0]['user'], justify_top)
            worksheet.write(next_row, 2, messages[0]['message'])
            
            next_row += 1
            reply_count = 1
            for reply in messages[1:]:
                worksheet.write(next_row, 0, messages[0]['timestamp'])
                worksheet.write(next_row, 1, reply['user'], justify_top)
                worksheet.write(next_row, 2, None)
                worksheet.write(next_row, 3, reply['message'])
                next_row += 1
                reply_count += 1
            
    workbook.close()




