import imaplib, re, email, datetime, settings
from email.header import decode_header

class Attachment:
    def __init__(self, filename, content):
        self.filename = filename
        self.__content = content

    def save_as(self, filepath):
        file = open(filepath, 'wb')
        file.write(self.__content)
        file.close()

class Message:
    def __init__(self, address_from, subject, date, plain_text, html_text, attachments):
        self.address_from = address_from
        self.date = date
        self.subject = subject
        self.plain_text = plain_text
        self.html_text = html_text
        self.attachments = attachments

class MailReader:
    position = 0
    count = 0
    items = []

    def __init__(self, server, user, password, port, attachment_dir=None, search='ALL', content_types=[], allowed_extensions=[]):
        self.server = server
        self.user = user
        self.password = password
        self.port = port
        self.attachment_dir = attachment_dir
        self.search = search
        self.content_types = content_types
        self.allowed_extensions = allowed_extensions

    def open_connection(self):
        self.__init_imap()
        self.__init_items()
	
    def close_connection(self):
        self.imap.close()
        self.imap.logout()

    def has_next(self):
        return self.count is not 0 and self.position < self.count

    def get_next(self):
        self.position += 1

    def get_item(self):
        item = self.items[self.position]
        message = email.message_from_bytes(item[0][1])
        return self.__formate_response_message(message)

    def __init_imap(self):
        self.imap = imaplib.IMAP4_SSL(self.server, self.port)
        self.imap.login(self.user, self.password)
        self.imap.select('INBOX')

    def __init_items(self):
        result, data = self.imap.uid('search', None, self.search)
        for uid in data[0].split():
            result, item = self.imap.uid('fetch', uid, '(RFC822)')
            self.items.append(item)
        self.count = len(self.items)

    def __formate_response_message(self, message):
        address_from = re.search(r'(?<=\<)(.*)(?=\>)', message['From']).group(0)
        subject = self.__decode_message_part(message['Subject'])
        date = self.__datetime_from_tuple(message['Date'])
        plain_text, html_text, attachments = self.__parse_content(message)
        response = Message(address_from, subject, date, plain_text, html_text, attachments)
        return response

    def __decode_message_part(self, message_part):
        decoded_part = decode_header(message_part)[0]
        if decoded_part[1]:
            return decoded_part[0].decode(decoded_part[1])
        else:
            return message_part

    def __datetime_from_tuple(self, date_tuple):
        date = email.utils.parsedate(date_tuple)
        return datetime.datetime(*date[0:6])

    def __parse_content(self, message):
        html_text, plain_text, attachments = '', '', []
        for part in message.walk():
            plain_text, html_text = self.__get_text_by_content_type(part)
            attachment = self.__get_attachment(part)
            if attachment is not None:
                attachments.append(attachment)
        return plain_text, html_text, attachments

    def __get_text_by_content_type(self, part):
        content_type = part.get_content_type()
        plain_text, html_text = '', ''
        if content_type in self.content_types:
            if content_type == 'text/plain':
                plain_text = part.get_payload(decode=True).decode(part.get_content_charset())
            elif content_type == 'text/html':
                html_text = part.get_payload(decode=True).decode(part.get_content_charset())
            return plain_text, html_text
        return None, None
                

    def __get_attachment(self, message_part):
        if message_part.get_content_maintype() is not 'multipart' and message_part.get('Content-Disposition') is not None:
            filename = self.__decode_message_part(message_part.get_filename())
            extension = filename.split('.')[-1]
            if bool(filename) and extension in self.allowed_extensions:
                attachment = Attachment(filename, message_part.get_payload(decode=True))
                return attachment

########################################################

def write_message(file, message, counter):
    file.write("###################################\n")
    file.write("Message: %d\n" % (counter,))
    file.write("###################################\n")
    file.write("%s\n" % (message.address_from,))
    file.write("%s\n" % (message.subject,))
    file.write("%s\n" % (message.date.strftime('%d.%m.%Y %H:%M:%S'),))
    file.write("%s\n" % (message.plain_text,))

def save_attachments(message):
    if message.attachments:
        for attachment in message.attachments:
            attachment.save_as(settings.ATTACHMENT_DIR + attachment.filename)

if __name__ == '__main__':

    iterator = MailReader(
        server=settings.CONNECTION['server'],
        user=settings.CONNECTION['user'],
        password=settings.CONNECTION['password'],
        port=settings.CONNECTION['port']
    )
    iterator.open_connection()

    iterator.attachment_dir = settings.ATTACHMENT_DIR
    iterator.content_types = settings.CONTENT_TYPES
    iterator.allowed_extensions = settings.ALLOWED_EXTENSIONS

    counter = 0
    file = open('emails.log', 'w')

    while iterator.has_next():
        counter += 1
        message = iterator.get_item()
        
        write_message(file, message, counter)
        save_attachments(message)
        
        iterator.get_next()

    iterator.close_connection()
    file.write("***************************\n")
    file.write("Total: %d\n" % (counter,))
    file.write("***************************\n\n")
    file.close()