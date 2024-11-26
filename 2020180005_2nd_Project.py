import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import email
from email.header import decode_header
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from PIL import Image, ImageTk

# GUI 설정
window = Tk()
window.title("My Project")
window.geometry('800x600')
window.resizable(True, True)

image = None
label = None
file_name = None  # 파일 이름을 저장할 전역 변수 추가


def open_image():
    global image, label, file_name  # file_name 전역 변수 사용 선언

    file_name = filedialog.askopenfilename(
        title='Select Images',
        filetypes=(
            ("Png Images", '*.png'),
            ("Jpg Images", '*.jpg'),
            ("All Images", '*.*')
        )
    )

    if file_name:
        if label is not None:
            label.destroy()

        image = Image.open(file_name)
        photo = ImageTk.PhotoImage(image)
        label = Label(window, image=photo)
        label.image = photo
        label.pack()
        window.geometry(f'{image.size[0]}x{image.size[1]}')


# 이메일 전송 함수
def send_email(subject, body, sender_email, receiver_email, password, image_path):
    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = receiver_email

    # 텍스트 본문 추가
    message.attach(MIMEText(body, 'plain'))

    # 이미지 파일 첨부
    if image_path:
        with open(image_path, 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-Disposition', 'attachment', filename=image_path)
            message.attach(img)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print("이메일을 성공적으로 보냈습니다.")


# 이메일 전송 인터페이스
def email_interface():
    # 이메일 정보 입력을 위한 새 창 생성
    email_window = Toplevel(window)
    email_window.title("이메일 보내기")

    # 이메일 입력 필드
    Label(email_window, text="제목:").grid(row=0, column=0)
    subject_entry = Entry(email_window, width=50)
    subject_entry.grid(row=0, column=1)

    Label(email_window, text="본문:").grid(row=1, column=0)
    body_entry = Text(email_window, width=50, height=10)
    body_entry.grid(row=1, column=1)

    Label(email_window, text="받는이 이메일:").grid(row=2, column=0)
    receiver_email_entry = Entry(email_window, width=50)
    receiver_email_entry.grid(row=2, column=1)

    # 이메일 전송 버튼
    send_button = Button(email_window, text="이메일 보내기", command=lambda: send_email(
        subject_entry.get(),
        body_entry.get("1.0", END),
        "",  # 발신자 이메일
        receiver_email_entry.get(),
        "",  # 앱 비밀번호
        file_name  # 이미지 파일 경로
    ))
    send_button.grid(row=3, column=1)


# IMAP 서버에 로그인하고 특정 키워드가 포함된 메일을 삭제하는 함수
def delete_emails_with_keyword(imap_url, username, password, keyword):
    # IMAP 서버에 연결
    mail = imaplib.IMAP4_SSL(imap_url)
    mail.login(username, password)

    # 받은편지함 선택
    mail.select('inbox')

    # 키워드를 UTF-8로 인코딩
    keyword_encoded = keyword.encode('UTF-8')

    # 키워드를 포함하는 메일 검색
    status, messages = mail.search(None, 'BODY', keyword_encoded)

    # 검색된 메일이 있는 경우
    if status == 'OK':
        for num in messages[0].split():
            # 메일 데이터 가져오기
            typ, data = mail.fetch(num, '(RFC822)')
            # 메일 파싱
            msg = email.message_from_bytes(data[0][1])
            # 메일 제목 디코딩
            subject = decode_header(msg['subject'])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            print(f'메일 삭제: {subject}')
            # 메일 삭제
            mail.store(num, '+FLAGS', '\\Deleted')

        # 삭제된 메일을 실제로 제거
        mail.expunge()

    # 로그아웃
    mail.logout()

def delete_email_interface():
    # 이메일 삭제를 위한 새 창 생성
    delete_window = Toplevel(window)
    delete_window.title("이메일 삭제하기")

    # 키워드 입력 필드
    Label(delete_window, text="삭제할 키워드:").grid(row=0, column=0)
    keyword_entry = Entry(delete_window, width=50)
    keyword_entry.grid(row=0, column=1)

    # 이메일 삭제 버튼
    delete_button = Button(delete_window, text="이메일 삭제하기", command=lambda: delete_emails_with_keyword(
        'imap.gmail.com',
        "",  # 사용자 이메일
        "",  # 앱 비밀번호
        keyword_entry.get()  # 삭제할 키워드
    ))
    delete_button.grid(row=1, column=1)

# 이메일 삭제 인터페이스

# 메뉴 설정
top_menu = Menu()

menu_File = Menu(top_menu, tearoff=0)
menu_File.add_command(label='Open Picture', accelerator='Ctrl+O', command=open_image)
menu_File.add_command(label='Send Email', command=email_interface)  # 이메일 보내기 메뉴 추가
menu_File.add_command(label='Delete Email', command=delete_email_interface)  # 이메일 삭제 메뉴 추가
menu_File.add_separator()
menu_File.add_command(label='Quit', command=window.quit)
top_menu.add_cascade(label='E-Mail', menu=menu_File)
window.config(menu=top_menu)

window.mainloop()
