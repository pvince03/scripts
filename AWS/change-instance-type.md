INSTANCE_ID="i-0a25ba0e201a0db3f"
NEW_TYPE="c6i.large"

# Stop instance
aws ec2 stop-instances --instance-ids $INSTANCE_ID
aws ec2 wait instance-stopped --instance-ids $INSTANCE_ID

# Modify instance type
aws ec2 modify-instance-attribute --instance-id $INSTANCE_ID --instance-type "{\"Value\": \"$NEW_TYPE\"}"

# Start instance
aws ec2 start-instances --instance-ids $INSTANCE_ID
