import discord
import openpyxl
from os import name


client = discord.Client()
token = "NzY4NDM5ODEyNDg2NzI1NjMz.X5AfWA.RJx2Q8dwvJdAGAdcpxka16SoMfg"

@client.event
async def on_ready():
    print(client.user.id)
    print("ready")


@client.event
async def on_message(message):

    if message.content.startswith("-질문"):
        embed = discord.Embed(title="자주묻는질문",color=0xEF2A59)
        embed.set_author(name="Alice World")
        embed.add_field(name="ㅤ", value="ㅤ", inline=False)
        embed.add_field(name="|명령어 목록", value="ㅤ아래에 명령어로 검색하면 궁금증이 해결됩니다.", inline=False)
        embed.add_field(name="-몬파", value="몬파 보상", inline=True)
        embed.add_field(name="-핫타임", value="핫타임 시간", inline=True)
        embed.add_field(name="-테섭", value="테스트 보상", inline=True)
        embed.add_field(name="-홍보", value="본섭 유지관련", inline=True)
        embed.add_field(name="-본섭", value="본섭 오픈관련", inline=True)
        await message.channel.send(embed=embed)  

    if message.content.startswith("-몬파"):
        embed = discord.Embed(title="몬스터 파크 요일별 보상",color=0x00aaaa)
        embed.set_author(name="Alice World")
        embed.add_field(name="ㅤ", value="ㅤ", inline=False)   
        embed.add_field(name="월요일", value="몬스터 파크 보상", inline=False)
        embed.add_field(name="화요일", value="아케인 심볼 선택권 X 3", inline=False)
        embed.add_field(name="수요일", value="에디셔널 큐브 X 1", inline=False)
        embed.add_field(name="목요일", value="영원한 환생의 불꽃 X 1", inline=False)
        embed.add_field(name="금요일", value="메소 주머니 (1000만~ 1억) X 1", inline=False)
        embed.add_field(name="토요일", value="경험치 2배 쿠폰 (30분) X 1", inline=False)
        embed.add_field(name="일요일", value="월~토 보상 중 랜덤 획득", inline=False)
        await message.channel.send(embed=embed)

    if message.content.startswith("-핫타임"):
        embed = discord.Embed(color=0x00aaaa)
        embed.add_field(name="Alice World", value="금일 PM 9:00 입니다")
        await message.channel.send(embed=embed)

    if message.content.startswith("-테섭"):
        embed = discord.Embed(title="테스트 서버 보상",color=0x00aaaa)
        embed.set_author(name="Alice World")
        embed.add_field(name="ㅤ", value="ㅤ", inline=False)       
        embed.add_field(name="|보상 내용", value="https://discord.gg/HZnVQK9", inline=False)
        embed.add_field(name="|훈장", value="Alice 선발대 지급", inline=True)
        embed.add_field(name="|계정 보상", value="계정 경험치에 따라 [테스터코인]이 지급 됩니다.", inline=True)
        await message.channel.send(embed=embed)

    if message.content.startswith("-홍보"):
        embed = discord.Embed(title="사전 홍보",color=0x00aaaa)
        embed.set_author(name="Alice World")
        embed.add_field(name="ㅤ", value="ㅤ", inline=False)       
        embed.add_field(name="|내용 확인", value="https://discord.gg/JjaBG9b", inline=False)
        embed.add_field(name="|지급", value="홍보확인후 테스트 서버에 즉시지급 됩니다.", inline=False)
        embed.add_field(name="|사전 예약 ", value="본섭오픈 3일전부터 본섭 기준 홍보포인트 지급이됩니다.", inline=False)  
        embed.add_field(name="|유지관련", value="테스트 서버에서 포인트를 사용하여도 본섭오픈시 사용전 포인트가 지급됩니다.", inline=True)        
        await message.channel.send(embed=embed)

    if message.content.startswith("-본섭"):
        embed = discord.Embed(title="D-DAYㅤ- 10",color=0x00aaaa)
        embed.set_author(name="Alice World")
        embed.add_field(name="ㅤ", value="ㅤ", inline=False)
        embed.add_field(name="|사전 예약 ", value="ㅤ본섭오픈 3일전부터 시작됩니다.", inline=False)        
        embed.add_field(name="|오픈일", value="ㅤ2020.11.1 예정입니다.", inline=True)      
        await message.channel.send(embed=embed)

    if message.content.startswith("!DM"):
        author = message.guild.get_member(int(message.content[4:22]))
        msg = message.content[23:]
        await author.send(msg)

    if message.content.startswith("!채금"):
        author = message.guild.get_member(int(message.content[4:22]))
        role = discord.utils.get(message.guild.roles, name="채금")
        await author.add_roles(role)

    if message.content.startswith("!노채금"):
        author = message.guild.get_member(int(message.content[5:23]))
        role = discord.utils.get(message.guild.roles, name="채금")
        role1 = discord.utils.get(message.guild.roles, name="앨리스 유저")
        await author.remove_roles(role)
        await author.add_roles(role1)    

    if message.content.startswith("!경고"):
        author = message.guild.get_member(int(message.content[4:22]))
        file = openpyxl.load_workbook("경고.xlsx")
        sheet = file.active
        i = 1
        while True:
            if sheet["A" + str(i)]. value == str(author.id):
                sheet["B" + str(i)].value = int(sheet["B" + str(i)].value) + 1
                file.save("경고.xlsx")
                if sheet["B" + str(i)].value == 2:
                    await message.guild.ban(author)
                    await message.channel.send("```[" + str(author.display_name) + "] 님 ALICE WORLD에서 추방 됩니다.```")
                else:
                    await message.channel.send("```[" + str(author.display_name) + "] 님 경고 1회 입니다.```")
                    await message.channel.send("```누적 2회시 ALICE WORLD에서 추방 됩니다.```")
                break
            if sheet["A" + str(i)]. value == None:
                sheet["A" + str(i)]. value = str(author.id)
                sheet["B" + str(i)]. value = 1
                file.save("경고.xlsx")
                await message.channel.send("```[" + str(author.display_name) + "] 님 경고 1회 입니다.```")
                await message.channel.send("```누적 2회시 ALICE WORLD에서 추방 됩니다.```")
                break
            i += 1

client.run(token)
