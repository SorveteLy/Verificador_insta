import requests, openpyxl, time
from bs4 import BeautifulSoup
wb = openpyxl.load_workbook(r"./Contas.xlsx")  
ws = wb.active
login = ws['S2']
senha = ws['U2']
Contas = ws['Q2']
numero_de_contas = Contas.value + 1 #numero que ta no Excel
numero = 2
Existe = 0
Nao_existe = 0
class color:
   PURPLE = '\033[95m'
   GREEN = '\033[92m'
   BOLD = '\033[1m'
   CWHITE  = '\33[37m'
   WARNING = '\033[31m'
inicio = time.time()
print(color.GREEN + "[+] " + color.CWHITE + "--" * 20 + color.GREEN + " [+]")
try:
    while (numero < numero_de_contas):
                c1 = ws['A{}'.format(numero)]
                request = requests.get('https://www.instagram.com/{}'.format(c1.value))
                if request.status_code == 200:
                    content = request.content              
                    source = BeautifulSoup(content, 'html.parser')
                    seg = 0
                    fot = 0
                    pes = 0
                    total = 0
                    try:     
                        description = source.find("meta", {"property": "og:description"}).get("content")
                        info_list = description.split("-")[0]
                        followers = info_list[0:info_list.index("Followers")]
                        info_list = info_list.replace(followers + "Followers, ", "")
                        following = info_list[0:info_list.index("Following")]
                        info_list = info_list.replace(following + "Following, ", "")
                        posts = info_list[0:info_list.index("Posts")]
                        results = {"followers": followers, "following": following, "posts": posts}
                        try:
                                print(color.GREEN + "[+] " + color.CWHITE + "Nome De Usuario:" + color.GREEN + " {}".format(c1.value))
                                #print(color.GREEN + "[+] " + color.CWHITE + "Seguidores:" + color.GREEN + "      {}".format(followers))
                                #print(color.GREEN + "[+] " + color.CWHITE + "Seguindo:" + color.GREEN + "        {}".format(following))
                                #print(color.GREEN + "[+] " + color.CWHITE + "Publicaçoes:" + color.GREEN + "     {}".format(posts))
                                #print(color.GREEN + "[+] " + color.CWHITE + "--" * 20 + color.GREEN + " [+]")
                                ws['I{}'.format(numero)] = "{}".format(posts)
                                ws['E{}'.format(numero)] = "{}".format(followers)               
                                ws['G{}'.format(numero)] = "{}".format(following)
                                ws['C{}'.format(numero)] = "Existente"
                                try:
                                    following = following.replace(",","")
                                except:
                                    exit
                                try:
                                    following = following.replace(".","")
                                except:
                                    exit
                                if int(followers) >= 15 and int(posts) >= 4 and int(following) > 0 :
                                    total += 1
                                if int(followers) >= 50 and int(posts) >= 6 and int(following) > 30 :
                                    total += 1
                                if int(followers) >= 100 and int(posts) >= 12 and int(following) > 30 :
                                    total += 1
                                if total == 1 :
                                    ws['m{}'.format(numero)] = "baixa"
                                elif total == 2 :
                                    ws['m{}'.format(numero)] = "media"
                                elif total == 3 :
                                    ws['m{}'.format(numero)] = "alta"
                                else:
                                    ws['m{}'.format(numero)] = "invalido"
                                Existe = Existe + 1
                                wb.save(filename="Contas.xlsx")
                        except:
                                exit
                    except:
                        ws['I{}'.format(numero)] = "inexistente"
                        ws['E{}'.format(numero)] = "inexistente"
                        ws['G{}'.format(numero)] = "inexistente"
                        ws['C{}'.format(numero)] = "inexistente"
                        ws['M{}'.format(numero)] = "inexistente"
                        print(color.GREEN + "[+]" + color.CWHITE + " Nome De Usuario:" + color.WARNING + " {}".format(c1.value))
                        #print(color.GREEN + "[+] " + color.CWHITE + "--" * 20 + color.GREEN + " [+]")
                        Nao_existe = Nao_existe + 1
                        wb.save(filename="Contas.xlsx")
                        exit
                else:
                    raise Exception(color.GREEN + "[+]" + color.CWHITE + " Nome De Usuario Invalido:" + color.WARNING + " {}".format(c1.value))
                numero = numero + 1
except:
        print("error")
        exit
fim = time.time()
print("\n")
print( color.GREEN + "[============================================]")
print( color.GREEN + "[+] " + color.CWHITE + "Tempo de Duraçao = {}".format(fim - inicio))
print( color.GREEN + "[+] " + color.CWHITE + "Total de Contas = {}".format(Existe + Nao_existe))
print( color.GREEN + "[+] " + color.CWHITE + "Contas Existentes = {}".format(Existe))
print( color.GREEN + "[+] " + color.CWHITE + "Contas Inexistente = {}".format(Nao_existe))
print( color.GREEN + "[===========================================]")
wb.save(filename="Contas.xlsx")