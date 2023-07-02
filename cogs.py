from discord.ext import commands
from discord import application_command
from registro import Regitrador
import logging
import discord

guildlist = ["1004735784403878019"]



class PontosTrabalho(commands.Cog):
    def __init__(self,client) -> None:
        self.client = client
        

        print(f"{__class__.__name__} inicialized!")
    
    @application_command(name="entrar",description="emtra no trabalho",guild_ids=guildlist)
    async def entrar(self,ctx):
        try:
            author = ctx.author
            register = Regitrador(f"{author}")
            
            register.EntradaRegistro()
            return await ctx.response.send_message(f"``` Usuario : {author} Entrou No trabalho. ```")
        except Exception as err:
            logging.error(f"Error in /entrar of {ctx.author} : {err}")
        

    @application_command(name="sair",description="sai do trabalho",guild_ids=guildlist)
    async def sair (self,ctx):
        
        try:
            author = ctx.author
            register = Regitrador(f"{author}")
            register.SaidaRegistro()
            return await ctx.response.send_message(f"```Usuario : {author} Saiu do Trabalho .\nHoras Trabalhadas:\n{register.HorasTrabalhadas()}  -30 minutos de pausa```")
        except Exception as err:
            logging.error(f"Error in /sair of {ctx.author} : {err}")
            
    @application_command(name="tabela_de_trabalho",description="pega a sua tabela de trabalho",guild_ids=guildlist)
    async def tabela_de_trabalho (self,ctx):
        try:
            author = ctx.author
            
            await ctx.response.send_message("Aqui esta sua Tabela:")
            with open(f"{author}.xlsx","rb") as fb:
                await ctx.send(file=discord.File(fb,f"{author}.xlsx"))
        except Exception as err:
            logging.error(f"Error in /tabela_de_trabalho of {ctx.author} : {err}")

def setup(client):
    cogs = [PontosTrabalho(client)]
    for x in cogs:
        client.add_cog(x)