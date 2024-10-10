# Instruções

1. **Renomeie o Arquivo de Configuração**:
   - Renomeie o arquivo `.env.sample` para `.env`.

2. **Preencha as Credenciais**:
   - Abra o arquivo `.env` e preencha as credenciais necessárias. As variáveis de ambiente esperadas são:

   ```
   TENANT_ID=<seu_tenant_id> 
   CLIENT_ID=<seu_client_id> 
   CLIENT_SECRET=<seu_client_secret>
   ```

Substitua `<seu_tenant_id>`, `<seu_client_id>`, e `<seu_client_secret>` pelos valores apropriados que você obteve durante o registro do aplicativo no Azure.

3. **Instale as Dependências**:
- Execute o comando a seguir para instalar todas as dependências do projeto:

```bash
npm install
