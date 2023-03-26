
  <?php

namespace App\Providers;

use App\Exceptions\CouldNotGetToken;
use App\Exceptions\CouldNotReachService;
use App\Exceptions\CouldNotSendMail;
use App\Transports\MicrosoftGraphMailTransport;
use GuzzleHttp\Exception\BadResponseException;
use GuzzleHttp\Exception\ConnectException;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\ServiceProvider;
use Microsoft\Graph\Graph;

class MicrosoftGraphServiceProvider extends ServiceProvider
{
    /**
     * Bootstrap services.
     *
     * @return void
     * @throws CouldNotSendMail
     */
    public function boot(): void
    {
        /**
         * Mail Transport
         */
        $this->app->get('mail.manager')->extend('microsoftgraph', function(array $config = []){

            if(!isset($config['transport']) || !$this->app['config']->get('mail.from.address', false)){
                throw CouldNotSendMail::invalidConfig();
            }

            return new MicrosoftGraphMailTransport();
        });
    }

    /**
     * Register services.
     *
     * @return void
     */
    public function register(): void
    {
        $this->app->bind(Graph::class, function (){

            return (new Graph())->setAccessToken(self::getAccessToken());

        });

    }
    /**
     * Get AccessToken
     */
    public static function getAccessToken(){

        return Cache::remember('microsoftgraph-accesstoken', 45, function (){

            try {

                $config = config('microsoftgraph');

                $guzzle = new \GuzzleHttp\Client();
                $response = $guzzle->post("https://login.microsoftonline.com/{$config['tenant']}/oauth2/token?api-version=1.0", [
                    'form_params' => [
                        'client_id' => $config['clientid'],
                        'client_secret' => $config['clientsecret'],
                        'resource' => 'https://graph.microsoft.com/',
                        'grant_type' => 'client_credentials',
                    ],
                ]);

                $response = json_decode((string)$response->getBody());

                return $response->access_token;

            } catch (BadResponseException $exception){

                // The endpoint responded with 4XX or 5XX error
                $response = json_decode((string)$exception->getResponse()->getBody());

                throw CouldNotGetToken::serviceRespondedWithError($response->error, $response->error_description);

            } catch (ConnectException $exception){

                // A connection error (DNS, timeout, ...) occurred
                throw CouldNotReachService::networkError();

            } catch (\Exception $Exception){

                // An unknown error occurred
                throw CouldNotReachService::unknownError();
            }
        });
    }

    /**
     * Get the services provided by the provider.
     *
     * @return array
     */
    public function provides(){
        return ['microsoftgraph'];
    }
}

   