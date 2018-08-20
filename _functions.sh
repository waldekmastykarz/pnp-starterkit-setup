isError() {
  res=$(echo $1 | jq -r '.message')
  if [[ -z "$res" || "$res" = "null" ]]; then return 1; else return 0; fi
}

success() {
  echo -e "\033[32m$1\033[0m"
}

warning() {
  echo -e "\033[33m$1\033[0m"
}

error() {
  echo -e "\033[31m$1\033[0m"
}

errorMessage() {
  msg=$(echo "$1" | jq -r ".message")
  error "$msg"
}

# $1 string with key-value pairs
# $2 name of the property for which to retrieve value
getPropertyValue() {
  echo "$1" | grep -o "$2:\"[^\"]\\+" | cut -d"\"" -f2
}