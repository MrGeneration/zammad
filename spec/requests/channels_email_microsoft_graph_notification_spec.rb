# Copyright (C) 2012-2026 Zammad Foundation, https://zammad-foundation.org/

require 'rails_helper'

RSpec.describe 'Microsoft Graph Email Notification', aggregate_failures: true, authenticated_as: :admin, type: :request do
  let(:admin) { create(:admin) }

  describe 'GET /api/v1/channels_email' do
    it 'offers Microsoft Graph for the notification channel only' do
      get '/api/v1/channels_email', as: :json

      expect(json_response['channel_driver']).to include(
        'email'              => include('outbound' => not_include('microsoft_graph_outbound')),
        'email_notification' => include('smtp', 'sendmail', 'microsoft_graph_outbound'),
      )
    end
  end

  describe 'POST /api/v1/channels_email_notification (microsoft_graph_outbound)' do
    it 'returns a redirect to the OAuth URL' do
      post '/api/v1/channels_email_notification', params: { adapter: 'microsoft_graph_outbound' }, as: :json

      expect(response).to have_http_status(:ok)
      expect(json_response).to include(
        'result' => 'redirect',
        'url'    => include('external_credentials/microsoft_graph/link_account', 'notification=true'),
      )
    end

    context 'with shared mailbox' do
      it 'passes shared_mailbox in the redirect URL' do
        post '/api/v1/channels_email_notification', params: {
          adapter: 'microsoft_graph_outbound',
          options: { shared_mailbox: 'shared@example.com' },
        }, as: :json

        expect(json_response['url']).to include('shared_mailbox=shared%40example.com')
      end
    end
  end

  describe 'GET /api/v1/external_credentials/microsoft_graph/link_account (notification)' do
    before do
      # The OAuth flow keeps its state in the session, which only exists for a signed-in browser.
      authenticated_as(admin, via: :browser)

      create(:external_credential, name: 'microsoft_graph', credentials: {
               client_id:     'test_client_id',
               client_secret: 'test_client_secret',
               client_tenant: 'common',
             })
    end

    it 'stores notification flag in session and redirects to Microsoft OAuth' do
      get '/api/v1/external_credentials/microsoft_graph/link_account', params: { notification: 'true' }

      expect(response).to have_http_status(:found)
      expect(response.headers['Location']).to include('login.microsoftonline.com')
    end

    it 'requests permission to send only' do
      get '/api/v1/external_credentials/microsoft_graph/link_account', params: { notification: 'true' }

      expect(response.headers['Location'])
        .to include('mail.send')
        .and not_include('mail.readwrite')
    end

    context 'when running as online service' do
      before { Setting.set('system_online_service', true) }

      it 'is forbidden' do
        get '/api/v1/external_credentials/microsoft_graph/link_account', params: { notification: 'true' }

        expect(response).to have_http_status(:forbidden)
      end
    end
  end

  describe 'GET /api/v1/external_credentials/microsoft_graph/callback (notification)' do
    let(:state_token) { SecureRandom.urlsafe_base64 }

    let(:token_response) do
      {
        access_token:  'new_access_token',
        refresh_token: 'new_refresh_token',
        expires_in:    3600,
        scope:         'offline_access openid profile email mail.send',
        token_type:    'Bearer',
        id_token:      'fake.id.token',
      }
    end

    let(:shared_mailbox) { nil }
    let(:probe_result)   { { result: 'ok' } }

    let(:graph_notification_channel) do
      Channel.where(area: 'Email::Notification').find do |ch|
        ch.options.dig(:outbound, :adapter) == 'microsoft_graph_outbound'
      end
    end

    before do
      # The OAuth flow keeps its state in the session, which only exists for a signed-in browser.
      authenticated_as(admin, via: :browser)

      create(:external_credential, name: 'microsoft_graph', credentials: {
               client_id:     'test_client_id',
               client_secret: 'test_client_secret',
               client_tenant: 'common',
             })

      allow(ExternalCredential).to receive(:request_account_to_link).and_return(
        request_token: state_token,
        authorize_url: 'https://login.microsoftonline.com/dummy',
      )

      allow(ExternalCredential::MicrosoftGraph).to receive_messages(
        authorize_tokens: token_response,
        user_info:        { preferred_username: 'user@example.com' },
      )

      allow(EmailHelper::Probe).to receive(:outbound).and_return(probe_result)

      # Populate session via link_account so the real Rails session store is used
      get '/api/v1/external_credentials/microsoft_graph/link_account', params: {
        notification:   'true',
        shared_mailbox: shared_mailbox,
      }
    end

    it 'exchanges code for tokens and updates the notification channel' do
      get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: state_token }

      expect(response).to have_http_status(:found)
      expect(response.headers['Location']).to include('#channels/email')

      channel = Channel.where(area: 'Email::Notification').find do |ch|
        adapter = ch.options.dig(:outbound, :adapter) || ch.options.dig('outbound', 'adapter')
        adapter == 'microsoft_graph_outbound'
      end

      expect(channel).to have_attributes(
        active:  true,
        options: include(
          outbound: include(
            adapter: 'microsoft_graph_outbound',
            options: include(user: 'user@example.com'),
          ),
          auth:     include(
            provider:     'microsoft_graph',
            access_token: 'new_access_token',
          ),
        ),
      )
    end

    context 'with shared mailbox' do
      let(:shared_mailbox) { 'shared@example.com' }

      it 'stores shared mailbox in outbound options' do
        get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: state_token }

        channel = Channel.where(area: 'Email::Notification').find do |ch|
          adapter = ch.options.dig(:outbound, :adapter) || ch.options.dig('outbound', 'adapter')
          adapter == 'microsoft_graph_outbound'
        end

        expect(channel.options.dig(:outbound, :options, :shared_mailbox)).to eq('shared@example.com')
      end
    end

    it 'verifies the delivery with the notification sender before switching' do
      get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: state_token }

      expect(EmailHelper::Probe).to have_received(:outbound).with(
        {
          adapter: 'microsoft_graph_outbound',
          options: { user: 'user@example.com', password: 'new_access_token' },
        },
        Setting.get('notification_sender'),
      )
    end

    context 'when the mailbox does not accept the notification sender' do
      let(:probe_result) { { result: 'invalid', message: 'Microsoft Graph API: send as denied (ErrorSendAsDenied)' } }

      it 'does not switch the notification channel and reports the reason' do
        expect { get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: state_token } }
          .not_to change { Channel.where(area: 'Email::Notification', active: true).pluck(:id) }

        expect(response).to have_http_status(:found)
        expect(response.headers['Location']).to include('#channels/microsoft_graph/error/notification_verification_failed/param/', 'ErrorSendAsDenied')
        expect(graph_notification_channel).to have_attributes(active: false)
      end
    end

    context 'when admin consent is pending' do
      it 'redirects to the admin consent information' do
        get '/api/v1/external_credentials/microsoft_graph/callback', params: { error: 'access_denied', error_description: 'AADSTS65004: User declined to consent to access the app.' }

        expect(response).to have_http_status(:found)
        expect(response.headers['Location']).to end_with('#channels/microsoft_graph/error/AADSTS65004')
        expect(graph_notification_channel).to have_attributes(active: false)
      end
    end

    context 'when running as online service' do
      it 'is forbidden' do
        Setting.set('system_online_service', true)

        get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: state_token }

        expect(response).to have_http_status(:forbidden)
        expect(graph_notification_channel).to have_attributes(active: false)
      end
    end

    context 'with invalid state token' do
      it 'returns unprocessable content' do
        get '/api/v1/external_credentials/microsoft_graph/callback', params: { code: 'auth_code', state: 'wrong_state' }

        expect(response).to have_http_status(:unprocessable_content)
      end
    end
  end
end
