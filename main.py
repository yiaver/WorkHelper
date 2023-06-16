import os
from dotenv import load_dotenv
from discord.ext import commands

client = commands.Bot(command_prefix=".")

@client.event
async def on_ready():
    print(f"logado como {client.user}.")

@client.event
async def on_member_join():
    pass

@client.event
async def on_member_remove():
    pass

if __name__ == "__main__":
    load_dotenv()
    
    cog_files = ["cogs"]
    for cog_file in cog_files:
        client.load_extension(cog_file)
        print(f"{cog_file} Loaded!")
    
    client.run(os.getenv("TOKEN"))