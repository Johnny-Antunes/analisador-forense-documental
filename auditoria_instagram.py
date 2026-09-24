from datetime import datetime, timezone
import json
import instaloader

def coletar_perfil_instagram(username: str):
    bot = instaloader.Instaloader()
    
    print(f"[*] Consultando perfil público: @{username}...")
    try:
        perfil = instaloader.Profile.from_username(bot.context, username)
        
        # Estrutura com carimbo de tempo para auditoria
        dados_auditoria = {
            "metadata_coleta": {
                "alvo": username,
                "timestamp_utc": datetime.now(timezone.utc).isoformat(),
                "modo_coleta": "anonimo_publico"
            },
            "perfil": {
                "user_id": perfil.userid,
                "username": perfil.username,
                "nome_completo": perfil.full_name,
                "biografia": perfil.biography,
                "link_externo": perfil.external_url,
                "seguidores": perfil.followers,
                "seguindo": perfil.followees,
                "conta_comercial": perfil.is_business_account,
                "categoria_comercial": perfil.business_category_name,
                "conta_verificada": perfil.is_verified,
                "foto_perfil_url": perfil.profile_pic_url
            }
        }
        
        # Gera nome de arquivo padronizado
        timestamp_str = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"ig_audit_{username}_{timestamp_str}.json"
        
        with open(nome_arquivo, "w", encoding="utf-8") as f:
            json.dump(dados_auditoria, f, ensure_ascii=False, indent=4)
            
        print(f"[+] Coleta concluída com sucesso!")
        print(f"[+] Evidência salva em: {nome_arquivo}")
        
    except instaloader.exceptions.ProfileNotExistsException:
        print(f"[-] Erro: O perfil @{username} não foi encontrado.")
    except instaloader.exceptions.LoginRequiredException:
        print("[-] Erro: O Instagram exigiu autenticação para visualizar este perfil.")
    except Exception as e:
        print(f"[-] Falha na requisição: {e}")

if __name__ == "__main__":
    alvo = input("Digite o @ do usuário (sem o @): ").strip()
    if alvo:
        coletar_perfil_instagram(alvo)
