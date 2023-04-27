# Settings that configure how the script behaves
# The most common ones to edit are below

# Where we will output the results
# "." is the project home directory
WORKING_DIRECTORY="."

# The name that we will export our data under
# Any existing file of this name will be deleted
EMAIL_DATA_DUMP="email-data.xlsx"

# The Name of the shared outlook inbox we want to walk 
INBOX_NAME="Business Response"

# Maximum number of emails that we will process
# Set to -1 if you want to process the entire folder
BREAK_AFTER_X_MAILS=-1

# Flush the cache to disk after X emails then continue
# It means we still have (most) information even if there is an error
FLUSH_AFTER_X_MAILS=20

#####

# Most of the time you will not need to edit these settings
LOG_FILE="ear.log"

