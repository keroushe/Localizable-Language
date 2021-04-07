import re
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

str = '"home_ms"="123";'

# lst = re.findall('"(\s|\S){1,}?";{0,1}', str)
lst = re.findall('(\s|\S){1,}', str)
print lst

pattern = re.compile('"(\s|\S){1,}?";{0,1}')
f = re.finditer(pattern, str)
print f

#"setting_notification_close"="The push notification of AKASO GO is not turned on. Go to \"Settings\" > \"Notifications\" and find \"AKASO GO\" to enable notification.";

for match in f:
    s = match.start()
    e = match.end()
    print 'String match "%s" at %d:%d' % (str[s:e], s, e)