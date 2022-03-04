create or replace package body pkg_Send_mails as

procedure pr_send_mails(p_sender varchar2,
                        p_recipents varchar2,
                        p_subject varchar2,
                        p_html_body clob,
                        p_path_file varchar2)
is
BEGIN
  null;
EXECUTE IMMEDIATE 'ALTER SESSION SET smtp_out_server = ''127.0.0.1''';
 IF p_path_file = '*' THEN
  UTL_MAIL.send(sender     => p_sender,
                recipients => p_recipents,
                MIME_TYPE  =>'text/html',
                cc         => null,
                bcc        => null,
                subject    => p_subject,
                message    => p_html_body);
  ELSE

  UTL_MAIL.send_attach_varchar2(sender     => p_sender,
                recipients => p_recipents,
                MIME_TYPE  =>'text/html',
                cc         => null,
                bcc        => null,
                subject    => p_subject,
                message    => p_html_body,
                attachment   => 'The is the contents of the attachment.',
                att_filename => p_path_file);
 END IF;
END pr_send_mails;
end pkg_Send_mails;
/
