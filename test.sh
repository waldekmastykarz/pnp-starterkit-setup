foo() {
  site=$1
  list=$2
  echo "${@:3}"
}

foo site 'list' --Title abc --stuff def --more asd