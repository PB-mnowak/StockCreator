from tokenize import Token
import pytest

import sys
from os import path

target_dir = path.abspath(r"C:\Users\m.nowak\Desktop\Python\_PT_stock_creator_local\src")
sys.path.append(target_dir)
Token = __import__('classes').Token

@pytest.fixture()
def token_fixture():
    token = Token()
    token.token = None
    token.user = "LG_User7@purebiologics.com"
    token.password = "LG_user7"
    token.test_mode = True

def test_token_init():
    token = Token()
    assert token
    
def test_test_mode_true():
    token = Token()
    token.test_mode = True
    assert token.test_mode == True
    
def test_test_mode_false():
    token = Token()
    assert token.test_mode == False
    
def test_token_call():
    token = Token()
    token.token = "test_token"
    assert token() == "test_token"