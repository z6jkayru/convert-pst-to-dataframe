import pypff
import pandas as pd

pst = pypff.file()
pst.open("CONVERTED_1.pst")

root = pst.get_root_folder()

def getMB(bytes):
  return round(bytes/1048576, 2)

def parse_folder(base):
    messages = []
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            messages += parse_folder(folder)
        for message in folder.sub_messages:
            total_attachment_size_bytes = 0
            str_attachments = ""
            count_attachments = message.number_of_attachments
            if (count_attachments > 0):
                for index in range(0, count_attachments):
                    filename = message.get_attachment(index).get_name()
                    str_attachments += filename + ", "
                    attach_size = message.get_attachment(index).get_size()
                    total_attachment_size_bytes += attach_size
                    # content = (message.get_attachment(0).read_buffer(attach_size)).decode('ascii', errors="ignore")
            body = message.plain_text_body
            text = body.decode()
            messages.append({
              "subject": message.subject,
              "sender": message.sender_name,
              "header": message.transport_headers,
              "body": text,
              "creation_time": message.creation_time,
              "submit_time": message.client_submit_time,
              "delivery_time": message.delivery_time,
              "attachment_count": count_attachments,
              "total_attachment_size_mb": getMB(total_attachment_size_bytes),
              "attachments": str_attachments[:-2],
              "datetime": message.client_submit_time
            })
    return messages

messages = parse_folder(root)

df = pd.DataFrame(messages)
df['datetime'] = df['datetime'].dt.tz_localize(tz='UTC')
df['datetime'] = df['datetime'].dt.tz_convert(tz='America/Bogota')
df = df.sort_values(by = "datetime", ascending = False)
df['datetime'] = df['datetime'].astype(str)

filename = "mails_exported.xlsx"

df.to_excel(filename)