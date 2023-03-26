
class MicrosoftGraphMailTransport extends AbstractTransport {
  
    protected mixed $graph;

  
    public function __construct()
    {
        parent::__construct();
        $this->graph = App::make(Graph::class);
    }

    protected function doSend(SentMessage $message): void {

        $from = $message->getEnvelope()->getSender()->getAddress();

        $this->graph->createRequest("POST", "/users/{$from}/sendmail")->attachBody($this->getBody($message))->execute();
    }

    protected function getBody($message): array
    {
        $email = MessageConverter::toEmail($message->getOriginalMessage());

        return array_filter([
            'message' => [
                'subject' => $email->getSubject(),
                'sender' => $this->toRecipientCollection($email->getFrom())[0],
                'from' => $this->toRecipientCollection($email->getFrom())[0],
                'replyTo' => $this->toRecipientCollection($email->getReplyTo()),
                'toRecipients' => $this->toRecipientCollection($email->getTo()),
                'ccRecipients' => $this->toRecipientCollection($email->getCc()),
                'bccRecipients' => $this->toRecipientCollection($email->getBcc()),
                'importance' => $email->getPriority() === 3 ? 'Normal' : 'High',
                'body' => $this->getContent($email),
                'attachments' => $this->toAttachmentCollection($email->getAttachments()),
            ]
        ]);
    }



    protected function toRecipientCollection(array|string $recipients)
    {
        $collection = [];

        if(!$recipients){
            return $collection;
        }

        if(is_string($recipients)){

            $collection[] = [
                'emailAddress' => [
                    'name' => null,
                    'address' => $recipients->getAddress(),
                ],
            ];

            return $collection;
        }

        foreach($recipients as $address){
            $collection[] = [
                'emailAddress' => [
                    'name' => $address->getName(),
                    'address' => $address->getAddress(),
                ],
            ];
        }

        return $collection;
    }

    /**
     * @param Email $email
     * @return array
     */
    private function getContent(Email $email): array {

        if (!is_null($email->getHtmlBody())) {
            $content = [
                'contentType' => 'html',
                'content' => $email->getHtmlBody(),
            ];
        }

        if (!is_null($email->getTextBody())) {
            $content = [
                'contentType' => 'text',
                'content' => $email->getTextBody(),
            ];
        }

        return $content;
    }

    /**
     * Transforms given SwiftMailer children into
     * Microsoft Graph attachment collection
     * @param $attachments
     * @return array
     */
    protected function toAttachmentCollection($attachments)
    {
        $collection = [];

        foreach($attachments as $attachment){
            /** @var DataPart $attachment */

            $collection[] = [
                'name' => $attachment->getFilename(),
                'contentId' => $attachment->getContentId(),
                'contentType' => $attachment->getContentType(),
                'contentBytes' => base64_encode($attachment->getBody()),
                'size' => strlen($attachment->getBody()),
                '@odata.type' => '#microsoft.graph.fileAttachment',
                'isInline' => false,
            ];

        }

        return $collection;
    }


    /**
     * Get the string representation of the transport.
     *
     * @return string
     */
    public function __toString(): string{
        return 'microsoftgraph';
    }
}