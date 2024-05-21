# MailSender ![icons8-email-sent-64](https://github.com/adavid1/MailSender/assets/36786512/093b00cd-5bbd-4f51-8a7b-4dcffc09fe0a)


MailSender is a C# application developed during my internship at Silverlight Research in London. The purpose of this application is to facilitate the mass mailing of experts to schedule appointments with them. The application provides a user-friendly interface to manage email templates, configure mail body content, and send emails in bulk.

## Features

- Paste Dropbox file links
- Configure mail body content
- Select region for targeted mailing
- Send bulk emails with a single click

## User Interface

### Main Interface

![mailsender](https://github.com/adavid1/MailSender/assets/36786512/ea6f8edb-917a-4eca-8111-7679e9aaec3d)

- **Paste**: Button to paste the link of the Dropbox file.
- **Mail Body**: Button to configure the mail body content.
- **Region**: Dropdown to select the region for targeted mailing.
- **Send**: Button to send the emails.

### Mail Body Configuration

![mailsender2](https://github.com/adavid1/MailSender/assets/36786512/a23cd983-5694-495c-88c4-cc42cb8ea53b)

- **Mail's subject**: Text field to enter the subject of the email.
- **Mail body**: Text area to write or paste the body of the email.
- **Greetings language**: Dropdown to select the language of the greetings.
- **Body template**: Dropdown to choose the template for the email body.
- **Signature**: Button to browse and add a signature file.

## Prerequisites

Make sure you have the following installed on your system:

- .NET Framework
- Visual Studio

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/adavid1/MailSender.git
   cd MailSender
   ```

2. Open the Project in Visual Studio
Launch Visual Studio and open the MailSender.sln solution file.

3.Build the Solution
Build the solution to restore the NuGet packages and compile the project.

## Usage

1. **Run the Application**
   - Start the application by running it from Visual Studio or by executing the compiled `.exe` file.

2. **Configure Mail Body**
   - Click on the **Configure** button to set up the email subject and body. Select the appropriate greetings language and body template.

3. **Paste Dropbox File Link**
   - Use the **Paste** button to insert the link of the Dropbox file that you want to include in the email.

4. **Select Region**
   - Choose the desired region from the dropdown menu.

5. **Send Emails**
   - Click on the **Send** button to dispatch the emails to the selected recipients.

## Acknowledgements

- Silverlight Research for providing the opportunity to work on this project.
