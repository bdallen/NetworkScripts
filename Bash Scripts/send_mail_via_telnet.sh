#!/bin/bash
if [ $# = 4 ]; then
DOIT=1
WAIT_TIME=2
elif [ $# = 5 ]; then
DOIT=1
WAIT_TIME=$5
else
DOIT=0
echo "
Usage:
$0 host port from_email to_email [wait_time]
host        hostname or IP
port        e.g. 25
from_email  your email address
to_email    the e-mail address to send the test to
wait_time   optional amount of time to wait before sending the next command. default is 2sec
"
fi
 
if [ $DOIT = 1 ]; then
SERVER=$1
PORT=$2
FROM_MAIL=$3
TO_MAIL=$4
   (
        sleep $WAIT_TIME
        echo "EHLO btg_test"
        sleep $WAIT_TIME
        echo "MAIL FROM:$FROM_MAIL"
        sleep $WAIT_TIME
        echo "RCPT TO:$TO_MAIL"
        sleep $WAIT_TIME
        echo "DATA"
        sleep $WAIT_TIME
        echo "Subject:Test message"
        echo "From:$FROM_MAIL"
        echo "To:$TO_MAIL"
        echo " "
        echo "Hello."
        echo "This is a test message."
        echo "Bye."
        echo "."
        sleep $WAIT_TIME
        echo "QUIT"
   ) | telnet $SERVER $PORT
 
fi
