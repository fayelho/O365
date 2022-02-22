using Microsoft.Graph;
using System.Threading.Tasks;

namespace O365.Mail
{
    public interface IO365
    {
        Task SendMail(Message message);
    }
}
