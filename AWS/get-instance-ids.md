#must upload a names.txt file which has the list of servers

while read -r name; do
  name="${name%$'\r'}"
  id=$(aws ec2 describe-instances \
    --filters "Name=tag:Name,Values=$name" \
    --query "Reservations[].Instances[].InstanceId" \
    --output text)
  echo -e "$name\t$id"
done < names.txt