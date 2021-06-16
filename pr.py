import pyupbit as coin

access = "YAlNo1067caXX17XG9UjddxMzkDKgZPwtAJmfRwa"
secreat = "SQb5oZDsLfj5BEGnIMbLxMpqcqS5PK86FOTSK4aH"
account = coin.Upbit(access, secreat)

print(account.get_balance())