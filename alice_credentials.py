from pya3 import *

user_id = "AB18400"
api_key = "cjVq0mmo9XD6sWlGWZfIXVED1qJnd5KA3Ms7Ba4zx72O8bHpIthak094jnAfmm2LbPftfKY97xP5Ldy6HQoa5igcKXYR85KOV3pUpZCzx0TPrtjhXC2n3xYu"

def login():

	alice = Aliceblue(user_id= user_id,api_key= api_key)
	alice.get_session_id() # Get Session ID

	print("Auto Login Success")


	# alice.get_contract_master("MCX")
	# alice.get_contract_master("NFO")
	alice.get_contract_master("NSE")
	# alice.get_contract_master("BSE")
	# alice.get_contract_master("CDS")
	# alice.get_contract_master("BFO")
	# alice.get_contract_master("INDICES")
	return alice
