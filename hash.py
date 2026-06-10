import bcrypt

hash_value = b"$2b$10$6mpwZ.3TB2NNR1ypKPIvduG5X2fbuoqGKoIEaCGKpKzSC3AcaM8Am"
password = b"mat_khau_can_thu"

print(bcrypt.checkpw(password, hash_value))