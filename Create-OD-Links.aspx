<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, initial-scale=1.0">
    <title>Create OD Sync Link</title>
    <link rel='stylesheet'
          href='https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.2.3/css/bootstrap.min.css' />
    <!-- Vue -->
    <script src='https://cdnjs.cloudflare.com/ajax/libs/vue/2.6.10/vue.js'></script>
    <!-- Pnp -->
    <script src='https://cdnjs.cloudflare.com/ajax/libs/sp-pnp-js/3.0.10/pnp.min.js'></script>
    <script type="text/javascript">
        document.addEventListener('DOMContentLoaded', function () {
            new Vue({
                el: '#app',
                data: {
                    url: null,
                    site: null,
                    context: null,
                    web: null,
                    docs: null,
                    list: null,
                    email: null
                },
                methods: {
                    // Setup context
                    setupContext: async function () {
                        if (this.url) {
                            await $pnp.setup({ sp: { baseUrl: this.url } })
                            this.web = await $pnp.sp.web.get();
                            var lists_docs = await $pnp.sp.web.lists.get();
                            this.docs = lists_docs.filter(item => item.BaseType === 1 && item.Hidden === false)
                        }
                    }
                },
                filters: {
                },
                components: {
                },
                computed: {
                    link: function () {
                        if (this.web && this.list) {
                            return `odopen://sync/?siteId=${this.web.Id}&amp;webId=${this.web.Id}&amp;listId=${this.list.Id}&amp;webUrl=${this.web.Url}&userEmail=${this.email}`
                        }

                        return '#';
                    }
                },
                created: function () {
                }
            });
        }, false);
    </script>
</head>

<body>
    <div id="app"
         class="container">
        <div class="row py-3">
            <div class="col">
                <label>Enter site URL:</label>
                <input type="text"
                       class="form-control"
                       v-model="url" />
                <a class="btn btn-primary mt-3"
                   @click.stop.prevent="setupContext()">Setup context</a>
            </div>
        </div>
        <div class="row py-3"
             v-if="web && docs">
            <div class="col">
                <label>Select a library to share</label>
                <select class="form-control"
                        v-model="list">
                    <option v-for="doc in docs"
                            :value="doc">{{doc.Title}}</option>
                </select>
            </div>
        </div>

        <div class="row py-3"
             v-if="web && list">
            <div class="col">
                <label>Share to email:</label>
                <input type="email"
                       class="form-control"
                       v-model="email"
                       placeholder="Enter email" />
            </div>
        </div>

        <div class="row py-3"
             v-if="url && list && email">
            <div class="col">
                <h5>Link generated</h5>
                <a :href="link">{{link}}</span>
            </div>
        </div>
    </div>
</body>

</html>