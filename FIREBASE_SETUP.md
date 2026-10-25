# Firebase setup — IndimediaWifi

Project ID: `indimediawifi`

## One-time setup in Firebase Console

1. Open Firebase Console and select the **indimediawifi** project.
2. In **Build → Authentication → Sign-in method**, enable **Email/Password**.
3. In **Authentication → Users**, create the admin account manually. The app intentionally has no public sign-up page.
4. In **Build → Firestore Database**, create the database if it does not exist yet.
5. Deploy the repository's `firestore.rules` rules before using the admin app. They allow reads/writes only for signed-in Email/Password accounts. Do not use open rules such as `allow read, write: if true`.

## Deploy rules from a terminal

Install Firebase CLI and sign in with an account that has permission on this Firebase project, then run from the repository root:

```bash
firebase login
firebase use indimediawifi
firebase deploy --only firestore:rules
```

The repository includes `firebase.json` pointing to `firestore.rules`.

## App behavior

- Admin login uses Firebase Authentication (email + password).
- Customer and billing collections are read in real time from Firestore collections `customers` and `bills`.
- Create/edit/delete operations write to Firestore.
- Excel restore writes the imported records to Firestore; backup exports the current data loaded in the app.
- Demo records are **not** automatically inserted. Once both Firestore collections load and are empty, the dashboard offers **Isi data demo ke Firestore**. This action is optional and writes sample data to the live project.

## Important

The frontend code is committed to GitHub, but rules must be deployed to Firebase and the admin account must be created in Firebase Console before end-to-end use. Never put a service-account private key in this frontend repository.
