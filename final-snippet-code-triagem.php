
/**
 * Endpoint e integrações "store-decision" — versão completa e atualizada
 * - Centraliza lógica de horários, zonas, sessão e logging
 * - Mantém a mesma assinatura/contratos do snippet original
 * - Adiciona cutoffs por dia, fechamento da Central e conciliação por loja (Stripe/Pagar.me)
 */

/** ==============================
 *  CONSTANTES & MAPAS
 *  ============================== */
if (!defined('AFONSOS_STORE_CENTRAL'))  define('AFONSOS_STORE_CENTRAL',  'Central Distribuição (Sagrada Família)');
if (!defined('AFONSOS_STORE_BARREIRO')) define('AFONSOS_STORE_BARREIRO', 'Unidade Barreiro');
if (!defined('AFONSOS_STORE_SION'))     define('AFONSOS_STORE_SION',     'Unidade Sion');

if (!function_exists('afonsos_store_id_map')) {
    function afonsos_store_id_map(): array {
        return [
            AFONSOS_STORE_BARREIRO => '110727',
            AFONSOS_STORE_SION     => '127163',
            AFONSOS_STORE_CENTRAL  => '86261',
        ];
    }
}

if (!function_exists('afonsos_zone_maps')) {
    function afonsos_zone_maps(): array {
        return [
            'barreiro' => [24, 11, 13, 107, 10, 20, 132, 25, 59, 60, 131, 64, 66, 50, 31, 39, 133, 134, 135, 136],
            'sion'     => [114, 137, 12, 123, 124, 103, 101, 104, 21, 138, 139, 140, 141, 142, 143, 144],
        ];
    }
}

/** ==============================
 *  CATÁLOGO DE MÉTODOS DE PAGAMENTO
 *  ============================== */
if (!function_exists('afonsos_payment_catalog')) {
    function afonsos_payment_catalog(): array {
        return [
            // Stripe por loja
            ['id' => 'stripe',        'title' => 'Cartão de Crédito On-line'],
            ['id' => 'stripe_cc',     'title' => 'Cartão de Crédito On-line'],
            ['id' => 'eh_stripe_pay', 'title' => 'Cartão de Crédito On-line'],

            // Outros métodos
            ['id' => 'pagarme_custom_pix',     'title' => 'Pix'],
            ['id' => 'custom_729b8aa9fc227ff', 'title' => 'Cartão na Entrega'],
            ['id' => 'woo_payment_on_delivery','title' => 'Dinheiro na Entrega'],
            ['id' => 'custom_e876f567c151864', 'title' => 'Vale Alimentação'],
        ];
    }
}

/** ==============================
 *  CONTAS DE PAGAMENTO POR LOJA
 *  ============================== */
if (!function_exists('afonsos_payment_accounts')) {
    function afonsos_payment_accounts(): array {
        return [
            AFONSOS_STORE_CENTRAL  => ['stripe' => 'stripe',        'pagarme' => 'central'],
            AFONSOS_STORE_BARREIRO => ['stripe' => 'stripe_cc',     'pagarme' => 'barreiro'],
            AFONSOS_STORE_SION     => ['stripe' => 'eh_stripe_pay', 'pagarme' => 'sion'],
        ];
    }
}

/** ==============================
 *  LOGGING
 *  ============================== */
if (!function_exists('afonsos_write_log')) {
    function afonsos_write_log(string $filename, string $message, bool $rotate_daily = false): void {
        $log_dir  = WP_CONTENT_DIR;
        $log_file = trailingslashit($log_dir) . $filename;

        // Rotação diária simples (opcional)
        if ($rotate_daily && file_exists($log_file)) {
            $current_date  = wp_date('Y-m-d');
            $log_file_date = wp_date('Y-m-d', filemtime($log_file));
            if ($current_date !== $log_file_date) {
                @rename($log_file, $log_dir . '/' . str_replace('.log', '', $filename) . "-{$log_file_date}.log");
                // mantém últimos 30
                $arch = glob($log_dir . '/' . str_replace('.log', '', $filename) . '-*.log') ?: [];
                usort($arch, fn($a, $b) => filemtime($b) <=> filemtime($a));
                foreach (array_slice($arch, 30) as $old) { @unlink($old); }
            }
        }

        if (!file_exists($log_file)) {
            @file_put_contents($log_file, '');
            @chmod($log_file, 0664);
        }

        $timestamp = wp_date('Y-m-d H:i:s');
        $formatted = "=== [{$timestamp}] ===\n{$message}\n\n";
        file_put_contents($log_file, $formatted, FILE_APPEND | LOCK_EX);
    }
}

if (!function_exists('my_custom_log')) {
    function my_custom_log(string $message): void {
        afonsos_write_log('debug-final-store.log', $message, true);
    }
}

if (!function_exists('log_store_decision_error')) {
    function log_store_decision_error(string $message): void {
        afonsos_write_log('store-decision-errors.log', $message, false);
    }
}

/** ==============================
 *  HELPERS GERAIS
 *  ============================== */
if (!function_exists('afonsos_pickup_store_id')) {
    function afonsos_pickup_store_id(string $store_name): string {
        $map = afonsos_store_id_map();
        return $map[$store_name] ?? $map[AFONSOS_STORE_CENTRAL];
    }
}

if (!function_exists('afonsos_normalize_cep')) {
    function afonsos_normalize_cep(?string $cep): string {
        return $cep ? preg_replace('/\D/', '', $cep) : '';
    }
}

if (!function_exists('afonsos_ensure_wc_session')) {
    function afonsos_ensure_wc_session(): void {
        if (function_exists('WC') && WC() && is_null(WC()->session)) {
            WC()->session = new WC_Session_Handler();
            WC()->session->init();
        }
    }
}

if (!function_exists('afonsos_is_future_date')) {
    function afonsos_is_future_date(?string $scheduled_date, string $today_ymd): bool {
        if (!$scheduled_date) return false;
        try {
            $d1 = new DateTime($scheduled_date);
            $d0 = new DateTime($today_ymd);
            return $d1 > $d0;
        } catch (Exception $e) {
            my_custom_log("Erro ao parsear data '{$scheduled_date}': " . $e->getMessage());
            return false;
        }
    }
}

if (!function_exists('afonsos_time_between')) {
    function afonsos_time_between(string $now_hm, string $start_hm, string $end_hm): bool {
        return ($now_hm >= $start_hm && $now_hm <= $end_hm);
    }
}

/** ==============================
 *  CUTOFFS (loja física) & FECHAMENTO (Central)
 *  ============================== */
if (!function_exists('afonsos_store_cutoff_hm_by_weekday')) {
    function afonsos_store_cutoff_hm_by_weekday(): array {
        return [
            1 => '18:45', // seg
            2 => '18:45', // ter
            3 => '18:45', // qua
            4 => '18:45', // qui
            5 => '18:45', // sex
            6 => '17:45', // sáb
            7 => '13:45', // dom  (básico que funciona)
        ];
    }
}

if (!function_exists('afonsos_central_close_hm_by_weekday')) {
    function afonsos_central_close_hm_by_weekday(): array {
        return [
            1 => '20:00', // seg
            2 => '20:00',
            3 => '20:00',
            4 => '20:00',
            5 => '20:00',
            6 => '20:00', // sáb
            7 => '16:00', // dom
        ];
    }
}

/** ==============================
 *  LÓGICA DE JANELA (forçar Central)
 *  ============================== */
if (!function_exists('afonsos_force_central_window')) {
    function afonsos_force_central_window(bool $is_future_date): bool {
        if ($is_future_date) return false;

        $day = (int) wp_date('N');   // 1..7 (seg..dom)
        $now = wp_date('H:i');

        $store_cutoff = afonsos_store_cutoff_hm_by_weekday()[$day] ?? '18:45';
        $central_close = afonsos_central_close_hm_by_weekday()[$day] ?? '20:00';

        // Passou do cutoff da loja, mas ainda dentro do horário da Central → força Central
        return ($now > $store_cutoff && $now <= $central_close);
    }
}

if (!function_exists('afonsos_get_zone_id_by_cep')) {
    function afonsos_get_zone_id_by_cep(string $cep): ?int {
        if (!$cep) return null;
        $package = [
            'destination' => [
                'country'  => 'BR',
                'state'    => 'MG',
                'postcode' => $cep,
                'city'     => '',
                'address'  => '',
            ],
        ];
        try {
            $zone = WC_Shipping_Zones::get_zone_matching_package($package);
            return $zone ? (int) $zone->get_id() : null;
        } catch (Exception $e) {
            my_custom_log('Erro ao obter zona: ' . $e->getMessage());
            return null;
        }
    }
}

if (!function_exists('afonsos_store_from_zone')) {
    function afonsos_store_from_zone(?int $zone_id): string {
        $maps = afonsos_zone_maps();
        if ($zone_id && in_array($zone_id, $maps['barreiro'], true)) return AFONSOS_STORE_BARREIRO;
        if ($zone_id && in_array($zone_id, $maps['sion'],     true)) return AFONSOS_STORE_SION;
        return AFONSOS_STORE_CENTRAL;
    }
}

if (!function_exists('afonsos_effective_store')) {
    function afonsos_effective_store(string $store_final, bool $is_future_date, bool $force_central): string {
        // Pedido futuro mantém loja física; hoje, janela crítica força Central
        if ($is_future_date && in_array($store_final, [AFONSOS_STORE_BARREIRO, AFONSOS_STORE_SION], true)) {
            return $store_final;
        }
        if ($force_central && in_array($store_final, [AFONSOS_STORE_BARREIRO, AFONSOS_STORE_SION], true)) {
            return AFONSOS_STORE_CENTRAL;
        }
        return $store_final;
    }
}

if (!function_exists('afonsos_available_payment_methods')) {
    function afonsos_available_payment_methods(string $effective_store, string $shipping_method, int $current_day): array {
        // Hoje todas as lojas usam o mesmo catálogo; mantido para futuras regras
        $methods = afonsos_payment_catalog();

        // Exemplo de restrição futura (mantido do seu código; hoje desativado)
        $restrict_weekdays = false;
        if ($shipping_method === 'delivery'
            && in_array($effective_store, [AFONSOS_STORE_BARREIRO, AFONSOS_STORE_SION], true)
            && $current_day >= 1 && $current_day <= 4
            && $restrict_weekdays
        ) {
            $whitelist = ['stripe', 'pagarme_custom_pix'];
            $methods = array_values(array_filter($methods, fn($m) => in_array($m['id'], $whitelist, true)));
        }

        // Garante nome do Stripe padronizado
        foreach ($methods as &$m) {
            if ($m['id'] === 'stripe') $m['title'] = 'Cartão de Crédito On-line';
        }
        unset($m);

        return $methods;
    }
}

if (!function_exists('afonsos_store_decision_endpoint_url')) {
    function afonsos_store_decision_endpoint_url(): string {
        // Usa o próprio site por padrão
        return rest_url('custom/v1/store-decision');
    }
}

/** ==============================
 *  CONTEXTO DE ERROS APENAS DO ENDPOINT
 *  ============================== */
global $is_store_decision_endpoint;
$is_store_decision_endpoint = false;

set_error_handler(function($errno, $errstr, $errfile, $errline) {
    global $is_store_decision_endpoint;
    if (!$is_store_decision_endpoint) return false;
    if ($errno === E_DEPRECATED || $errno === E_USER_DEPRECATED) return true;
    $msg = "Erro não tratado no endpoint /store-decision\nTipo: $errno\nMensagem: $errstr\nArquivo: $errfile\nLinha: $errline";
    log_store_decision_error($msg);
    return false;
});

register_shutdown_function(function() {
    global $is_store_decision_endpoint;
    if (!$is_store_decision_endpoint) return;
    $error = error_get_last();
    if ($error && in_array($error['type'], [E_ERROR, E_PARSE, E_CORE_ERROR, E_COMPILE_ERROR], true)) {
        $msg = "Erro fatal detectado no endpoint /store-decision\nTipo: {$error['type']}\nMensagem: {$error['message']}\nArquivo: {$error['file']}\nLinha: {$error['line']}";
        log_store_decision_error($msg);
    }
});

/** ==============================
 *  ENDPOINT /store-decision
 *  ============================== */
add_action('rest_api_init', function() {
    my_custom_log('Registrando endpoint /store-decision');
    register_rest_route('custom/v1', '/store-decision', [
        'methods'             => 'POST',
        'callback'            => 'afonsos_calculate_store_decision',
        'permission_callback' => '__return_true',
    ]);
});

if (!function_exists('afonsos_calculate_store_decision')) {
    function afonsos_calculate_store_decision(WP_REST_Request $request) {
        global $is_store_decision_endpoint;
        $is_store_decision_endpoint = true;

        ini_set('memory_limit', '256M');
        @set_time_limit(60);

        // Garante WooCommerce
        if (!function_exists('WC') || !WC()) {
            my_custom_log('WooCommerce não está carregado. Tentando carregar...');
            @require_once ABSPATH . 'wp-content/plugins/woocommerce/woocommerce.php';
            if (!function_exists('WC') || !WC()) {
                my_custom_log('Falha ao carregar o WooCommerce.');
                $is_store_decision_endpoint = false;
                return new WP_Error('woocommerce_not_loaded', 'WooCommerce não pôde ser carregado', ['status' => 500]);
            }
        }
        afonsos_ensure_wc_session();

        try {
            my_custom_log('Endpoint /store-decision chamado');
            $params          = $request->get_json_params() ?: [];
            $cep             = afonsos_normalize_cep($params['cep']         ?? '');
            $shipping_method = sanitize_text_field($params['shipping_method'] ?? 'delivery');
            $pickup_store    = sanitize_text_field($params['pickup_store']    ?? '');
            $delivery_date   = sanitize_text_field($params['delivery_date']   ?? '');
            $pickup_date     = sanitize_text_field($params['pickup_date']     ?? '');

            if (!$cep) {
                my_custom_log('Erro: CEP não fornecido na requisição.');
                return new WP_Error('no_cep', 'CEP não fornecido', ['status' => 400]);
            }

            my_custom_log("CEP recebido: {$cep}");

            $current_day  = (int) wp_date('N');
            $current_time = wp_date('H:i');
            $current_date = wp_date('Y-m-d');

            // Data agendada?
            $scheduled_date = $delivery_date ?: $pickup_date;
            $is_future_date = afonsos_is_future_date($scheduled_date, $current_date);
            my_custom_log('Data agendada? ' . ($scheduled_date ? $scheduled_date : 'não') . ' | Future? ' . ($is_future_date ? 'true' : 'false'));

            // 1) Se for pickup e loja explicitamente escolhida
            if ($shipping_method === 'pickup' && $pickup_store) {
                $store_final     = $pickup_store;
                $pickup_store_id = afonsos_pickup_store_id($pickup_store);
                my_custom_log("Store Final por pickup: {$store_final} (pickup_store_id={$pickup_store_id})");
            } else {
                // 2) Delivery: decide por zona
                my_custom_log("Obtendo zona para CEP {$cep}");
                $zone_id         = afonsos_get_zone_id_by_cep($cep);
                $store_final     = afonsos_store_from_zone($zone_id);
                $pickup_store_id = afonsos_pickup_store_id($store_final);
                my_custom_log("Zona " . ($zone_id ?? 'N/A') . " => Store Final: {$store_final}");
            }

            // 3) Ajustes por janela crítica e/ou pedido futuro
            $force_central          = afonsos_force_central_window($is_future_date);
            $effective_store_final  = afonsos_effective_store($store_final, $is_future_date, $force_central);

            // 4) Métodos de pagamento (catálogo/base)
            $available_payment_methods = afonsos_available_payment_methods($effective_store_final, $shipping_method, $current_day);

            // 5) Contas de pagamento (conciliação)
            $accounts = afonsos_payment_accounts()[$effective_store_final] ?? ['stripe' => 'central', 'pagarme' => 'central'];

            // Persistência em sessão
            if (is_object(WC()->session)) {
                WC()->session->set('checkout_store_final',      $store_final);
                WC()->session->set('available_payment_methods', $available_payment_methods);
                WC()->session->set('effective_store_final',     $effective_store_final);
                WC()->session->set('is_future_date',            $is_future_date);
                WC()->session->set('payment_accounts',          $accounts);
            }

            $response = [
                'store_final'           => $store_final,
                'effective_store_final' => $effective_store_final,
                'pickup_store_id'       => $pickup_store_id,
                'payment_methods'       => $available_payment_methods,
                'payment_accounts'      => $accounts,
            ];

            my_custom_log(
                "ENDPOINT store-decision\n" .
                "CEP: {$cep}\nShipping: {$shipping_method}\nPickup Store: {$pickup_store}\n" .
                "Delivery Date: {$delivery_date}\nPickup Date: {$pickup_date}\n" .
                "Store Final: {$store_final}\nEffective: {$effective_store_final}\n" .
                "Pickup Store ID: {$pickup_store_id}\n" .
                'Gateways: ' . wp_json_encode($available_payment_methods) . "\n" .
                'Accounts: ' . wp_json_encode($accounts) . "\n" .
                "Now: {$current_time} (D{$current_day})\nFuture? " . ($is_future_date ? 'true' : 'false')
            );

            $is_store_decision_endpoint = false;
            return rest_ensure_response($response);

        } catch (Exception $e) {
            $msg = "Erro no endpoint /store-decision\nMensagem: {$e->getMessage()}\nArquivo: {$e->getFile()}\nLinha: {$e->getLine()}";
            my_custom_log($msg);
            log_store_decision_error($msg);

            // Fallback seguro: CENTRAL
            $store_final                = AFONSOS_STORE_CENTRAL;
            $pickup_store_id            = afonsos_pickup_store_id($store_final);
            $effective_store_final      = $store_final;
            $available_payment_methods  = afonsos_payment_catalog();
            $accounts                   = afonsos_payment_accounts()[$effective_store_final] ?? ['stripe' => 'central', 'pagarme' => 'central'];

            if (is_object(WC()->session)) {
                WC()->session->set('checkout_store_final',      $store_final);
                WC()->session->set('available_payment_methods', $available_payment_methods);
                WC()->session->set('effective_store_final',     $effective_store_final);
                WC()->session->set('is_future_date',            false);
                WC()->session->set('payment_accounts',          $accounts);
            }

            $response = [
                'store_final'           => $store_final,
                'effective_store_final' => $effective_store_final,
                'pickup_store_id'       => $pickup_store_id,
                'payment_methods'       => $available_payment_methods,
                'payment_accounts'      => $accounts,
            ];

            my_custom_log(
                "ENDPOINT store-decision (Fallback)\n" .
                "Store Final: {$store_final}\nEffective: {$effective_store_final}\n" .
                "Pickup Store ID: {$pickup_store_id}\n" .
                'Gateways: ' . wp_json_encode($available_payment_methods) . "\n" .
                'Accounts: ' . wp_json_encode($accounts)
            );

            $is_store_decision_endpoint = false;
            return rest_ensure_response($response);
        }
    }
}

/** ==============================
 *  Checkout -> chama o endpoint e grava sessão
 *  ============================== */
add_action('woocommerce_checkout_update_order_review', 'afonsos_update_store_final_and_payment_methods', 10);
if (!function_exists('afonsos_update_store_final_and_payment_methods')) {
    function afonsos_update_store_final_and_payment_methods($posted_data) {
        parse_str($posted_data, $data);

        // método de envio: pickup x delivery
        $shipping_method_id = '';
        if (!empty($data['shipping_method']) && is_array($data['shipping_method'])) {
            $shipping_method_id = (string) reset($data['shipping_method']);
            if (strpos($shipping_method_id, ':') !== false) {
                $shipping_method_id = explode(':', $shipping_method_id)[0];
            }
        }
        $shipping_method = (strpos($shipping_method_id, 'pickup') !== false) ? 'pickup' : 'delivery';

        // campos
        $pickup_store  = isset($data['shipping_pickup_stores']) ? sanitize_text_field($data['shipping_pickup_stores']) : '';
        $delivery_date = isset($data['delivery_date']) ? sanitize_text_field($data['delivery_date']) : '';
        $pickup_date   = isset($data['pickup_date'])   ? sanitize_text_field($data['pickup_date'])   : '';

        // CEP (várias fontes)
        $cep = '';
        if (!empty($data['billing_postcode'])) {
            $cep = afonsos_normalize_cep($data['billing_postcode']);
        } elseif (!empty($data['shipping_postcode'])) {
            $cep = afonsos_normalize_cep($data['shipping_postcode']);
        } elseif (function_exists('WC') && WC()->customer) {
            $cep = afonsos_normalize_cep(WC()->customer->get_billing_postcode());
        }

        $request_body = [
            'cep'             => $cep,
            'shipping_method' => $shipping_method,
            'pickup_store'    => $pickup_store,
            'delivery_date'   => $delivery_date,
            'pickup_date'     => $pickup_date,
        ];

        my_custom_log('Chamando /store-decision com: ' . wp_json_encode($request_body));

        $response = wp_remote_post(
            afonsos_store_decision_endpoint_url(),
            [
                'method'  => 'POST',
                'headers' => ['Content-Type' => 'application/json'],
                'body'    => wp_json_encode($request_body),
                'timeout' => 30,
            ]
        );

        if (is_wp_error($response)) {
            my_custom_log('Erro ao chamar o endpoint: ' . $response->get_error_message());
            return;
        }

        $body = json_decode(wp_remote_retrieve_body($response), true);
      if (isset($body['store_final'])) {
    afonsos_ensure_wc_session();
    if (is_object(WC()->session)) {
        WC()->session->set('checkout_store_final', $body['store_final']);
        WC()->session->set('available_payment_methods', $body['payment_methods'] ?? []);
        if (isset($body['effective_store_final'])) {
            WC()->session->set('effective_store_final', $body['effective_store_final']);
        }
        if (isset($body['is_future_date'])) {
            WC()->session->set('is_future_date', $body['is_future_date']);
        }
        if (isset($body['payment_accounts'])) {
            WC()->session->set('payment_accounts', $body['payment_accounts']);
        }
        WC()->session->set('shipping_method', $shipping_method); // Confirme que esta linha está presente
        my_custom_log('Sessão WC atualizada no update_order_review');
    }
} else {
            my_custom_log('Resposta inválida do endpoint: ' . wp_remote_retrieve_body($response));
        }

        my_custom_log(
            'Checkout: ' .
            "Shipping={$shipping_method} | Pickup={$pickup_store} | DelivDate={$delivery_date} | PickupDate={$pickup_date} | " .
            'StoreFinal=' . ($body['store_final'] ?? 'N/A') . ' | Effective=' . ($body['effective_store_final'] ?? 'N/A') .
            ' | Accounts=' . wp_json_encode($body['payment_accounts'] ?? [])
        );
    }
}



/** ==============================
 *  Filtra Stripe certo por loja
 *  ============================== */
add_filter('woocommerce_available_payment_gateways', function($gateways) {
    if (!is_checkout() || is_wc_endpoint_url('order-pay')) return $gateways;

    afonsos_ensure_wc_session();

    // Loja efetiva da sessão (ou Central no fallback)
    $effective_store_final = (WC()->session instanceof WC_Session)
        ? WC()->session->get('effective_store_final', AFONSOS_STORE_CENTRAL)
        : AFONSOS_STORE_CENTRAL;

    // Pega contas de pagamento definidas no mapa
    $accounts  = afonsos_payment_accounts();
    $stripe_id = $accounts[$effective_store_final]['stripe'] ?? 'stripe';

    // Remove todos os Stripes que não são da loja
    foreach (['stripe','stripe_cc','eh_stripe_pay'] as $sid) {
        if ($sid !== $stripe_id && isset($gateways[$sid])) {
            unset($gateways[$sid]);
        }
    }

    my_custom_log("Stripe Multi-Store → Loja efetiva: {$effective_store_final} | Stripe ativo: {$stripe_id}");
    my_custom_log("Gateways finais: " . wp_json_encode(array_keys($gateways)));

    return $gateways;
}, 150);


/** ==============================
 *  Filtragem de gateways no checkout
 *  ============================== */
add_filter('woocommerce_available_payment_gateways', 'afonsos_filter_payment_gateways_by_store', 100);
if (!function_exists('afonsos_filter_payment_gateways_by_store')) {
    function afonsos_filter_payment_gateways_by_store($gateways) {
        if (!is_checkout() || is_wc_endpoint_url('order-pay')) return $gateways;

        afonsos_ensure_wc_session();

        $gateway_ids = array_keys($gateways);
        my_custom_log('Gateways antes da filtragem: ' . wp_json_encode($gateway_ids));

        $available_payment_methods = is_object(WC()->session)
            ? (WC()->session->get('available_payment_methods', []) ?: [])
            : [];

        // Dados do checkout
        $posted       = WC()->checkout()->get_posted_data();
        $delivery_date= isset($posted['delivery_date']) ? sanitize_text_field($posted['delivery_date']) : '';
        $pickup_date  = isset($posted['pickup_date'])   ? sanitize_text_field($posted['pickup_date'])   : '';
        $store_final  = is_object(WC()->session)
            ? WC()->session->get('checkout_store_final', AFONSOS_STORE_CENTRAL)
            : AFONSOS_STORE_CENTRAL;

        // CEP
        $cep = '';
        if (!empty($posted['billing_postcode'])) {
            $cep = afonsos_normalize_cep($posted['billing_postcode']);
        } elseif (!empty($posted['shipping_postcode'])) {
            $cep = afonsos_normalize_cep($posted['shipping_postcode']);
        } elseif (function_exists('WC') && WC()->customer) {
            $cep = afonsos_normalize_cep(WC()->customer->get_billing_postcode());
        }

        // Future?
        $current_date   = wp_date('Y-m-d');
        $scheduled_date = $delivery_date ?: $pickup_date;
        $is_future_date = afonsos_is_future_date($scheduled_date, $current_date);

        // Recalcula effective_store_final com base no CEP (apenas para delivery)
$effective_store_final = $store_final;

// Descobre se é pickup ou delivery (salvo na sessão pelo endpoint)
$shipping_method = is_object(WC()->session)
    ? (WC()->session->get('shipping_method', 'delivery'))
    : 'delivery';

if ($shipping_method !== 'pickup' && $cep) {
    $zone_id               = afonsos_get_zone_id_by_cep($cep);
    $effective_store_final = afonsos_store_from_zone($zone_id);
} else {
    my_custom_log("Mantendo effective_store_final='{$store_final}' (pickup ou sem CEP)");
}


        // Ajuste por janela crítica x futuro
        $force_central          = afonsos_force_central_window($is_future_date);
        $effective_store_final  = afonsos_effective_store($effective_store_final, $is_future_date, $force_central);

        // Salva em sessão para debug/integrações
        if (is_object(WC()->session)) {
            WC()->session->set('effective_store_final', $effective_store_final);
            WC()->session->set('is_future_date',        $is_future_date);
            my_custom_log('Sessão WC atualizada em filter_payment_gateways_by_store');
        }

        // Se o endpoint já trouxe a lista exata de métodos, aplica
        if (!empty($available_payment_methods)) {
            $allowed_ids = array_column($available_payment_methods, 'id');
            $filtered    = array_intersect_key($gateways, array_flip($allowed_ids));

            // garante stripe_multistore quando existir
            if (isset($gateways['stripe_multistore'])) {
                $filtered['stripe_multistore'] = $gateways['stripe_multistore'];
            }

            my_custom_log('Gateways finais (sessão): ' . wp_json_encode(array_keys($filtered)));
            return $filtered;
        }

        // Fallback: mantém todos + garante MultiStore
        if (isset($gateways['stripe_multistore'])) {
            $gateways['stripe_multistore'] = $gateways['stripe_multistore'];
            my_custom_log('Gateway stripe_multistore garantido no fallback');
        }

        my_custom_log('Gateways finais (fallback): ' . wp_json_encode(array_keys($gateways)));
        return $gateways;
    }
}

/** ==============================
 *  Debug em console (quando checkout atualiza)
 *  ============================== */
add_action('wp_footer', function() {
    if (is_checkout()) : ?>
        <script>
        jQuery(document.body).on('updated_checkout', function() {
            jQuery.post('<?php echo esc_url( admin_url('admin-ajax.php') ); ?>', {
                action: 'debug_effective_store_final'
            }, function(resp) {
                console.log('Stripe Multi-Store DEBUG → effective_store_final (sessão):', resp);
            });
        });
        </script>
    <?php endif;
});

/** ==============================
 *  Endpoint AJAX de debug
 *  ============================== */
add_action('wp_ajax_debug_effective_store_final', 'afonsos_debug_effective_store_final');
add_action('wp_ajax_nopriv_debug_effective_store_final', 'afonsos_debug_effective_store_final');
if (!function_exists('afonsos_debug_effective_store_final')) {
    function afonsos_debug_effective_store_final() {
        afonsos_ensure_wc_session();
        $store = is_object(WC()->session)
            ? (WC()->session->get('effective_store_final', '(desconhecido)') ?: '(desconhecido)')
            : '(sem sessão)';
        wp_send_json($store);
    }
}

/** ==============================
 *  Metadados no pedido
 *  ============================== */
add_action('woocommerce_checkout_create_order', 'afonsos_save_store_final_to_order', 20, 2);
if (!function_exists('afonsos_save_store_final_to_order')) {
    function afonsos_save_store_final_to_order($order, $data) {
        afonsos_ensure_wc_session();
        $order_id = $order->get_id() ?: '(pedido sem ID)';

        // Determina se é pickup
        $shipping_method = is_object(WC()->session)
            ? WC()->session->get('shipping_method', 'delivery')
            : 'delivery';

        // Obtém a loja escolhida para retirada, se houver
        $pickup_store = isset($_POST['shipping_pickup_stores']) ? sanitize_text_field((string) $_POST['shipping_pickup_stores']) : '';
        if (empty($pickup_store)) {
            $pickup_store = $order->get_meta('_shipping_pickup_stores', true);
        }

        // Define store_final e effective_store_final
        if ($shipping_method === 'pickup' && !empty($pickup_store)) {
            // Para retirada, usa a loja escolhida pelo cliente
            $store_final = $pickup_store;
            $effective_store_final = $pickup_store;
        } else {
            // Para entrega, usa os valores da sessão (baseados no CEP)
            $store_final = is_object(WC()->session)
                ? WC()->session->get('checkout_store_final', AFONSOS_STORE_CENTRAL)
                : AFONSOS_STORE_CENTRAL;
            $effective_store_final = is_object(WC()->session)
                ? WC()->session->get('effective_store_final', $store_final)
                : $store_final;
        }

        // Calcula o pickup_store_id com base no store_final
        $pickup_store_id = afonsos_pickup_store_id($store_final);

        // Outros valores da sessão
        $is_future_date = is_object(WC()->session)
            ? WC()->session->get('is_future_date', false)
            : false;
        $accounts = is_object(WC()->session)
            ? (WC()->session->get('payment_accounts', ['stripe' => 'central', 'pagarme' => 'central']))
            : ['stripe' => 'central', 'pagarme' => 'central'];

        // Salva metadados no pedido
        $order->update_meta_data('_store_final', $store_final);
        $order->update_meta_data('_effective_store_final', $effective_store_final);
        $order->update_meta_data('_is_future_date', $is_future_date ? 'yes' : 'no');
        $order->update_meta_data('_payment_account_stripe', $accounts['stripe'] ?? 'central');
        $order->update_meta_data('_payment_account_pagarme', $accounts['pagarme'] ?? 'central');

        // Salva metadados de retirada, se aplicável
        if (!empty($pickup_store)) {
            $order->update_meta_data('_shipping_pickup_stores', $pickup_store);
            $order->update_meta_data('_shipping_pickup_store_id', $pickup_store_id);
        }

        my_custom_log(
            "Checkout do site\n[PEDIDO #{$order_id}]\n" .
            "Salvando store_final: {$store_final}\n" .
            "Effective store: {$effective_store_final}\n" .
            "is_future_date: " . ($is_future_date ? 'yes' : 'no') . "\n" .
            "Stripe account: " . ($accounts['stripe'] ?? 'central') . "\n" .
            "Pagar.me account: " . ($accounts['pagarme'] ?? 'central') . "\n" .
            "Pickup Store: {$pickup_store}\n" .
            "Pickup Store ID: {$pickup_store_id}\n" .
            "Shipping Method: {$shipping_method}"
        );
    }
}