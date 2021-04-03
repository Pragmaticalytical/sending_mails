# Sending Emails with Python

<https://pragmaticalytical.com/website-coming-soon/tutorials/sending-emails-with-python/> 

When you start a new job in finance or analytics, the easiest and quickest way to impress your boss is to automate some aspect of the reporting workflow.

If the firm you work has a newly established or immature analytics/reporting department, the easiest part to automate of said workflow would usually be periodic email generation.

To help you score this quick win with your boss, and become the smart new kid on the office block everyone fears, here is a script that works with both Windows Outlook (if the server you use is Windows based) and SMTP servers (any kind of server).

I will create a tutorial on how to schedule Python jobs to run periodically on both Windows and Linux servers. For now though, these two scripts should help you send an email with attachments using Python. To schedule the job you will need to do a bit of further research. Here are some pointers to aid you:

If your computer/server has Windows installed, look at Windows Task Scheduler.
If your server is Linux based, look at Cron Jobs, or solutions like Apache Airflow if you are already advanced in your Python-fu.

To make this work on Windows, you will need a couple of special libraries installed on your machine. Install them by running:

`pip install -U pypiwin32 Jinja2`

If you are on Linux, it should work out of the box.
