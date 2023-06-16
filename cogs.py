from discord.ext import commands
from discord import application_command
from registro import Regitrador
import discord

guildlist = ["1004735784403878019"]

class PontosTrabalho(commands.Cog):
    def __init__(self,client) -> None:
        self.client = client
        

        print(f"{__class__.__name__} inicialized!")
    
    @application_command(name="entrar",description="emtra no trabalho",guild_ids=guildlist)
    async def entrar(self,ctx):
        author = ctx.author
        self.register = Regitrador(f"{author}")
        #role = discord.utils.get(ctx.guild.roles, name="Trabalhando")
        self.register.EntradaRegistro()
        await ctx.response.send_message(f"Usuario : {author} Entrou No trabalho. ")
        

    @application_command(name="sair",description="sai do trabalho",guild_ids=guildlist)
    async def sair (self,ctx):
        #role = discord.utils.get(ctx.guild.roles, name="Trabalhando")
        author = ctx.author
        
        self.register.SaidaRegistro()
        await ctx.response.send_message(f"Usuario : {author} Saiu do Trabalho .\nHoras Trabalhadas:\n *** {self.register.HorasTrabalhadas()} ***")

    @application_command(name="tabela_de_trabalho",description="pega a sua tabela de trabalho",guild_ids=guildlist)
    async def tabela_de_trabalho (self,ctx):
        author = ctx.author
        
        await ctx.response.send_message("Aqui esta sua Tabela:")
        with open(f"{author}.xlsx","rb") as fb:
            await ctx.send(file=discord.File(fb,f"{author}.xlsx"))

def setup(client):
    cogs = [PontosTrabalho(client)]
    for x in cogs:
        client.add_cog(x)