<?php

namespace App\Controllers;

use App\Models\UserModel;
use Microsoft\Graph\Graph;
use Beta\Microsoft\Graph\Model as BetaModel;

class Home extends BaseController
{
    public $db;
    public $session;

    public function __construct()
    {
        $this->db = \Config\Database::connect();
        $this->session = session();
    }

    public function index()
	{
	    $data['isLoggedIn'] = $this->session->get('logged_in');
		return view('welcome_message', $data);
	}

    /**
     * get user details by accessToken
     * and login if success
     * @throws \GuzzleHttp\Exception\GuzzleException
     */
	public function office365Auth()
    {
        $userModel = new UserModel();

        $accessToken = $_REQUEST['accessToken'];//$this->request->getVar('accessToken');
        if ( empty($accessToken) ) {
            echo json_encode(['status'=>'error', 'msg'=>'Invalid Request']);exit;
        }

        try {
            $graph = new Graph();
            $graph->setAccessToken($accessToken);

            $user = $graph->setApiVersion("beta")
                ->createRequest("GET", "/me")
                ->setReturnType(BetaModel\User::class)
                ->execute();

            $user_name = $user->getDisplayName();
            $user_email = $user->getUserPrincipalName();

            // adding or getting data from db
            $dbUser = $userModel->where('email', $user_email)->first();
            if ($dbUser) {
                // nothing
            } else { // insert to db
                $userModel->save([
                    'name'=>$user_name,
                    'email'=>$user_email
                ]);
                $dbUser = $userModel->where('email', $user_email)->first();
            }

            // storing data to session
            $session_data = [
                'user_id'       => $dbUser['id'],
                'user_name'     => $dbUser['name'],
                'user_email'    => $dbUser['email'],
                'logged_in'     => TRUE
            ];
            $this->session->set($session_data);

            echo json_encode(['status'=>'success', 'msg'=>'Login Success']);exit;

        } catch ( \Exception $e) {
            echo json_encode(['status'=>'error', 'msg'=>$e->getMessage()]);exit;
        }

    }

    public function logout()
    {
        $session = session();
        $session->destroy();
        return redirect()->to('/');
    }
}
