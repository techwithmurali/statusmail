create or replace package pkg_Send_mails  as
procedure pr_send_mails(p_recipents varchar2,
                        p_subject varchar2,
                        p_html_body clob,
                        p_path_file varchar2);
end pkg_Send_mails;
/					
